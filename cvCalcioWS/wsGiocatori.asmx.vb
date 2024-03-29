﻿Imports System.Web.Services
Imports System.ComponentModel
Imports System.IO
Imports System.Web.Hosting
Imports System.Diagnostics.Eventing.Reader
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_gioc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsGiocatori
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaCF(Cognome As String, Nome As String, Comune As String, DataNascita As String, Maschio As String) As String
		Dim cf As New CodiceFiscale
		Dim bMaschio As Boolean = IIf(Maschio = "S", True, False)
		Dim Ritorno As String = cf.CreaCodiceFiscale(Server.MapPath("."), Cognome, Nome, DataNascita, Comune, bMaschio)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ConvalidaFirma(idAnno As String, Squadra As String, idGiocatore As String, idGenitore As String, Mittente As String, QualeFirma As String) As String
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

				'Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				'If Ritorno<> "OK" Then
				'	Ok = False
				'End If

				If Ok Then
					Dim dataVal As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
					Sql = "Update GiocatoriFirme Set Validazione='" & dataVal & "' Where idGiocatore=" & idGiocatore & " And idGenitore=" & idGenitore & " And QualeFirma=" & QualeFirma
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						Dim c() As String = Squadra.Split("_")
						Dim Anno As String = Str(Val(c(0))).Trim
						Dim idSquadra2 As String = Str(Val(c(1))).Trim

						Sql = "Select * From [Generale].[dbo].[AggiornamentoWidgets] Where idSquadra=" & idSquadra2
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Sql = "Insert Into [Generale].[dbo].[AggiornamentoWidgets] Values (" & idSquadra2 & ", 'S')"
							Else
								Sql = "Update [Generale].[dbo].[AggiornamentoWidgets] Set AggiornaWidgets='S' Where idSquadra=" & idSquadra2
							End If
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							If Ritorno = "OK" Then
								Ritorno = "*"
							End If
						End If

						'Rec.Close()

						Sql = "Select A.Genitore" & idGenitore & ", MailGenitore" & idGenitore & ", TelefonoGenitore" & idGenitore & ", B.Maggiorenne From GiocatoriDettaglio A " &
							"Left Join Giocatori B On A.idGiocatore = B.idGiocatore Where A.idGiocatore=" & idGiocatore
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ritorno = StringaErrore & " Nessun genitore rilevato"
								Ok = False
							Else
								If idGenitore < 3 And idGenitore > -1 Then
									Dim Genitore As String = "" & Rec("Genitore" & idGenitore).Value
									Dim Mail As String = "" & Rec("MailGenitore" & idGenitore).Value
									Dim Telefono As String = "" & Rec("TelefonoGenitore" & idGenitore).Value
									Dim Maggiorenne As String = "" & Rec("Maggiorenne").Value

									Dim Attiva As String = "S"

									If Maggiorenne = "N" Then
										Attiva = "N"
									End If

									Sql = "Select * From GiocatoriMails Where idGiocatore=" & idGiocatore & " And Progressivo=" & idGenitore
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
									Else
										If Rec.Eof() Then
											Sql = "Insert Into GiocatoriMails Values (" & idGiocatore & ", " & idGenitore & ", '" & Mail.Replace("'", "''") & "', '" & Attiva & "')"
										Else
											Sql = "Update GiocatoriMails Set Mail='" & Mail.Replace("'", "''") & "' Where idGiocatore=" & idGiocatore & " And Progressivo=" & idGenitore
										End If
										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If
									End If

									If Ok Then
										Dim idGenitoreLetto As Integer = -1
										Dim GenitoreGiaEsisteComeUtente As Boolean = False
										Dim figliGiaPresenti As String = ""

										Sql = "Select * From [Generale].[dbo].[Utenti] Where EMail='" & Mail.Replace("'", "''") & "'"
										Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
										If TypeOf (Rec) Is String Then
											Ritorno = Rec
											Ok = False
										Else
											If Rec.Eof() Then
												GenitoreGiaEsisteComeUtente = False
											Else
												GenitoreGiaEsisteComeUtente = True
												idGenitoreLetto = "" & Rec("idUtente").Value
												figliGiaPresenti = "" & Rec("idGiocatore").Value
												If figliGiaPresenti = "-1" Then
													figliGiaPresenti = ""
												End If
												If Strings.Right(figliGiaPresenti, 1) <> ";" Then
													figliGiaPresenti = figliGiaPresenti & ";"
												End If
											End If
										End If
										'Rec.Close()

										If Ok Then
											If Not GenitoreGiaEsisteComeUtente Then
												If TipoDB = "SQLSERVER" Then
													Sql = "Select IsNull(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
												Else
													Sql = "Select Coalesce(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
												End If
												Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													'If Rec(0).Value Is DBNull.Value Then
													'	idGenitoreLetto = 1
													'Else
													idGenitoreLetto = Rec(0).Value
													'End If
												End If
											End If

											If Not Genitore.Contains(" ") Then
												Genitore = " " & Genitore
											End If
											Dim g() As String = Genitore.Split(" ")
											Dim s() As String = Squadra.Split("_")
											If s.Length > 0 Then
												Dim idSquadra As Integer = s(1)
												Dim pass As String = ""
												Dim conta As Integer = 0
												While Not pass.Contains(";")
													pass = generaPassRandom()
													conta += 1
													If conta > 20 Then
														Ritorno = StringaErrore & " Creazione password fallita"
														Ok = False
														Exit While
													End If
												End While

												If Ok Then
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
															"'N', " &
															"'" & stringaWidgets & "' " &
															")"
													Else
														figliGiaPresenti &= idGiocatore & ";"
														Sql = "Update [Generale].[dbo].[Utenti] Set " &
															"idGiocatore='" & figliGiaPresenti & "' " &
															"Where idUtente=" & idGenitoreLetto
													End If
													' COMMENTATO DIETRO RICHIESTA DI DONATO 14/09/2020
													'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
													'If Ritorno.Contains(StringaErrore) Then
													'	Ok = False
													'Else
													'	Dim m As New mail
													'	Dim Oggetto As String = "Nuovo utente inCalcio"
													'	Dim Body As String = ""
													'	Body &= "E' stato creato l'utente '" & Genitore.ToUpper & "'. <br />"
													'	Body &= "Per accedere al sito sarà possibile digitare la mail rilasciata alla segreteria in fase di iscrizione: " & Mail & "<br />"
													'	Body &= "La password valida per il solo primo accesso è: " & nuovaPass(0) & "<br /><br />"
													'	Dim ChiScrive As String = "notifiche@incalcio.cloud"

													'	Ritorno = m.SendEmail(Squadra, Mittente, Oggetto, Body, Mail, {""})
													'End If
													' COMMENTATO DIETRO RICHIESTA DI DONATO 14/09/2020
												End If
											Else
												Ok = False
												Ritorno = StringaErrore & " Problema nel ricavare i dati della società"
											End If
										End If
									End If
								Else
									If idGiocatore = -1 Then
										Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=-1 And QualeFirma=" & QualeFirma
										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If
									End If
								End If
							End If
						End If
					End If
				End If

				'If Ok Then
				'	Ritorno = "*"
				'	Sql = "Commit"
				'	Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				'Else
				'	Sql = "Rollback"
				'	Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				'End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaFirme(Squadra As String, idGiocatore As String, idGenitore As String) As String
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

				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)
				Dim NomeSquadra As String = ""

				Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						NomeSquadra = Rec("NomeSquadra").Value
					End If
				End If
				Rec.Close()

				If Ritorno = "" Then
					Dim gf As New GestioneFilesDirectory
					Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					gf = Nothing
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "\" Then
						Percorso &= "\"
					End If
					'Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_" & idGenitore & ".kgb"
					'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, idGenitore, "N") Then
					Try
						Sql = "Delete From immagini_firme Where id=" & idGiocatore & " And idGenitore=" & idGenitore
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & ": " & ex.Message
					End Try

					'Try
					'	Sql = "Delete From immagini_firme Where id=" & idGiocatore & " And Progressivo=" & idGenitore & " And Privacy='S'"
					'	Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					'Catch ex As Exception
					'	Ritorno = StringaErrore & ": " & ex.Message
					'End Try
					'Else
					'	Ritorno = StringaErrore & "Firma non esistente"
					'End If
				End If

				If Ritorno = "*" Or Ritorno = "OK" Then
					Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & idGenitore
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

					'Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & idGenitore + 100
					'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaFirma(Squadra As String, idGiocatore As String, idGenitore As String, QualeFirma As String) As String
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

				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)
				Dim NomeSquadra As String = ""

				Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						NomeSquadra = Rec("NomeSquadra").Value
					End If
				End If
				Rec.Close()

				If Ritorno = "" Then
					Dim gf As New GestioneFilesDirectory
					Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					gf = Nothing
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "\" Then
						Percorso &= "\"
					End If
					'Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_" & idGenitore & ".kgb"
					'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, idGenitore, "N") Then
					Try
						Sql = "Delete From immagini_firme Where id=" & idGiocatore & " And idGenitore=" & idGenitore & " And Progressivo=" & QualeFirma
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & ": " & ex.Message
					End Try

					'Try
					'	Sql = "Delete From immagini_firme Where id=" & idGiocatore & " And Progressivo=" & idGenitore & " And Privacy='S'"
					'	Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					'Catch ex As Exception
					'	Ritorno = StringaErrore & ": " & ex.Message
					'End Try
					'Else
					'	Ritorno = StringaErrore & "Firma non esistente"
					'End If
				End If

				If Ritorno = "*" Or Ritorno = "OK" Then
					Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & idGenitore & " And QualeFirma=" & QualeFirma
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

					'Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & idGenitore + 100
					'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiornaFirma(Squadra As String, ByVal idGiocatore As String, ByVal Genitore As String, QualeFirma As String, FirmaTablet As String, TipoUtente As String) As String
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

				If TipoUtente = "1" Then
					Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

					'If Privacy = "S" Then
					'	Genitore = Val(Genitore) + 100
					'End If

					Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
					Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore & " And QualeFirma=" & QualeFirma
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Rec.Close()
							Sql = "Insert Into GiocatoriFirme Values (" &
								" " & idGiocatore & ", " &
								" " & Genitore & ", " &
								"'" & Datella & "', " &
								"'', " &
								"'', " &
								" " & QualeFirma & " " &
								")"
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						Else
							Rec.Close()
						End If
					End If

					If FirmaTablet = "S" Then
						Sql = "Update GiocatoriFirme Set DataFirma='" & Datella & "', Validazione='" & Datella & "' Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore & " And QualeFirma=" & QualeFirma
					Else
						Sql = "Update GiocatoriFirme Set DataFirma='" & Datella & "' Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore & " And QualeFirma=" & QualeFirma
					End If
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

					If Ritorno = "*" Then
						Sql = "commit"
						Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					Else
						Sql = "rollback"
						Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					End If
				Else
					Ritorno = "*"
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ControllaFirma(Squadra As String, ByVal idGiocatore As String, ByVal Genitore As String, QualeFirma As String) As String
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

				Sql = "Select * From GiocatoriFirme A " &
					"Left Join Immagini_Firme B On A.idGiocatore = B.id And A.QualeFirma = B.Progressivo " &
					"Where A.idGiocatore=" & idGiocatore & " And A.idGenitore=" & Genitore & " And QualeFirma=" & QualeFirma
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						Dim Datella As String = Rec("DataFirma").Value

						If Trim(Datella) <> "" Then
							If Genitore <> 3 Then
								Ritorno = "Una firma è già stata inserita per il giocatore ed il genitore in data " & Datella & ";" & Rec("Dati").Value
							Else
								Ritorno = "Una firma è già stata richiesta per il giocatore in data " & Datella & ";" & Rec("Dati").Value
							End If
						Else
							Ritorno = "*"
						End If
					Else
						Ritorno = "*"
					End If
					Rec.Close()

					'Dim Giocatore As String = ""
					'Dim sGenitore As String = ""

					''If Ritorno = "*" Then
					'Sql = "Select * From Giocatori Where idGiocatore=" & idGiocatore
					'Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					'If TypeOf (Rec) Is String Then
					'	Ritorno = Rec
					'Else
					'	If Not Rec.Eof() Then
					'		Giocatore = Rec("Cognome").Value & " " & Rec("Nome").Value
					'	Else
					'		Ritorno = StringaErrore & " Giocatore non rilevato"
					'	End If
					'	Rec.Close()

					'	If Genitore <> 3 Then
					'		Sql = "Select * From GiocatoriDettaglio Where idGiocatore=" & idGiocatore
					'		Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					'		If TypeOf (Rec) Is String Then
					'			Ritorno = Rec
					'		Else
					'			Ritorno = ""
					'			If Not Rec.Eof() Then
					'				If Genitore <> 4 Then
					'					sGenitore = Rec("Genitore" & Genitore).Value
					'				Else
					'					sGenitore = Rec("Genitore1").Value
					'				End If
					'			Else
					'				Ritorno = StringaErrore & " Genitore non rilevato"
					'			End If
					'		End If
					'		Rec.Close()

					'		Ritorno &= Giocatore & ";" & sGenitore & ";"
					'	Else
					'		Ritorno &= Giocatore & ";;"
					'		'End If
					'	End If
					'End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RichiedeFirma(Squadra As String, ByVal idGiocatore As String, ByVal Genitore As String, Mittente As String, QualeFirma As String) As String
		' RichiedeFirma?Squadra= 0002_00160&idGiocatore=432&Genitore=1 
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
				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)

				ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "--------------------------------")
				ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Squadra: " & Squadra)
				ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Codice Squadra: " & codSquadra)
				ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Id Giocatore: " & idGiocatore)
				ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Genitore: " & Genitore)
				ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Mittente: " & Mittente)
				ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "QualeFirma: " & QualeFirma)
				'Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				'Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				'If Ritorno<> "OK" Then
				'	Sql = "rollback"
				'	Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				'	Return Ritorno
				'End If

				Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Err1: " & Ritorno)
				Else
					If Rec.Eof() Then
						Sql = "Insert Into GiocatoriFirme Values (" & idGiocatore & ", " & Genitore & ", '" & Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00") & "', '', '', " & QualeFirma & ")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					End If
				End If

				If Ritorno = "*" Or Ritorno = "" Or Ritorno = "OK" Then
					Ritorno = ""

					Sql = "Select NomeSquadra, Descrizione, iscrFirmaEntrambi From Anni Where idAnno = " & Anno
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Err2: " & Ritorno)
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessuna squadra rilevata"
							ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", Ritorno)
						Else
							Dim NomeSquadra As String = "" & Rec("NomeSquadra").Value
							Dim Descrizione As String = "" & Rec("Descrizione").Value
							Dim iscrFirmaEntrambi As String = "" & Rec("iscrFirmaEntrambi").Value
							Rec.Close()

							ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Nome Squadra: " & NomeSquadra)
							ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Descrizione: " & Descrizione)
							ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Iscrizione firma entrambi" & iscrFirmaEntrambi)

							Sql = "Select * From Giocatori Where idGiocatore = " & idGiocatore
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Err3: " & Ritorno)
							Else
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
									ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", Ritorno)
								Else
									Dim Nominativo As String = Rec("Cognome").Value & " " & Rec("Nome").Value
									Rec.Close()

									ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Nominativo: " & Nominativo)

									Sql = "Select MailGenitore1, MailGenitore2, B.Cognome + ' ' + B.Nome As Genitore3 , Genitore1, Genitore2, MailGenitore3, " &
										"B.Maggiorenne, GenitoriSeparati, AffidamentoCongiunto, idTutore " &
										"From GiocatoriDettaglio A " &
										"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
										"Where A.idGiocatore = " & idGiocatore
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", Ritorno)
									Else
										If Rec.Eof() Then
											Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", Ritorno)
										Else
											Dim EMail As String = ""
											Dim nomeGenitore As String = ""
											Dim Maggiorenne As String = "" & Rec("Maggiorenne").Value
											Dim GenitoriSeparati As String = "" & Rec("GenitoriSeparati").Value
											Dim AffidamentoCongiunto As String = "" & Rec("AffidamentoCongiunto").Value
											Dim idTutore As String = "" & Rec("idTutore").Value
											Dim ceGenitore1 As String = "" & Rec("Genitore1").Value
											Dim ceGenitore2 As String = "" & Rec("Genitore2").Value
											Dim Ok As Boolean = True
											Dim QualeGenitore As String = ""

											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Maggiorenne: " & Maggiorenne)
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Genitori Separati: " & GenitoriSeparati)
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Affidamento Congiunto: " & AffidamentoCongiunto)
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "idTutore: " & idTutore)
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "ceGenitore1: " & ceGenitore1)
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "ceGenitore2: " & ceGenitore2)

											If Genitore = "1" Then
												EMail = "" & Rec("MailGenitore1").Value
												nomeGenitore = "" & Rec("Genitore1").Value
												QualeGenitore = "1"
											Else
												If Genitore = "2" Then
													EMail = "" & Rec("MailGenitore2").Value
													nomeGenitore = "" & Rec("Genitore2").Value
													QualeGenitore = "2"
												Else
													If Genitore = "3" Then
														EMail = "" & Rec("MailGenitore3").Value
														nomeGenitore = "" & Rec("Genitore3").Value
														QualeGenitore = "3"
													Else
														If Maggiorenne = "S" Then
															EMail = "" & Rec("MailGenitore3").Value
															nomeGenitore = "" & Rec("Genitore3").Value
															QualeGenitore = "3"
														Else
															If GenitoriSeparati = "S" Then
																If AffidamentoCongiunto = "S" Then
																Else
																	If idTutore = "1" Then
																		EMail = "" & Rec("MailGenitore1").Value
																		nomeGenitore = "" & Rec("Genitore1").Value
																		QualeGenitore = "1"
																	Else
																		If idTutore = "2" Then
																			EMail = "" & Rec("MailGenitore2").Value
																			nomeGenitore = "" & Rec("Genitore2").Value
																			QualeGenitore = "2"
																		Else
																			Ok = False
																			Ritorno = StringaErrore & " Tutore non valido"
																		End If
																	End If
																End If
															Else
																If iscrFirmaEntrambi = "S" Then
																	If ceGenitore1 <> "" And ceGenitore2 <> "" Then
																		EMail = "" & Rec("MailGenitore1").Value
																		nomeGenitore = "" & Rec("Genitore1").Value
																		QualeGenitore = "1"
																	Else
																		Ok = False
																		Ritorno = StringaErrore & " Manca la presenza di uno o più genitori"
																	End If
																Else
																	If ceGenitore1 <> "" Then
																		EMail = "" & Rec("MailGenitore1").Value
																		nomeGenitore = "" & Rec("Genitore1").Value
																		QualeGenitore = "1"
																	Else
																		If ceGenitore2 <> "" Then
																			EMail = "" & Rec("MailGenitore2").Value
																			nomeGenitore = "" & Rec("Genitore2").Value
																			QualeGenitore = "2"
																		Else
																			Ok = False
																			Ritorno = StringaErrore & " Nessun genitore disponibile"
																		End If
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
											Rec.Close()

											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "EMail: " & EMail)
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Nome Genitore: " & nomeGenitore)
											ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Quale Genitore: " & QualeGenitore)

											If Ok Then
												Dim gf As New GestioneFilesDirectory
												Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PercorsoSito.txt")

												If Percorso = "" Then
													Ritorno = StringaErrore & " Nessun percorso sito rilevato"
													ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", Ritorno)
												Else
													Percorso = Percorso.Trim()
													If Strings.Right(Percorso, 1) <> "/" Then
														Percorso &= "/"
													End If

													Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
													Dim P() As String = PathAllegati.Split(";")
													If Strings.Right(P(0), 1) = "\" Then
														P(0) = Mid(P(0), 1, P(0).Length - 1)
													End If

													Dim gT1 As New GestioneTags(Server.MapPath("."))

													If Genitore = 4 Then
														' Gestione firma associato
														Dim m As New mail
														Dim Oggetto As String = "Richiesta Firma inCalcio"
														''Dim Body As String = gT1.EsegueMailAssociato(Server.MapPath("."), Squadra, idGiocatore, Anno, QualeGenitore, Privacy)
														'Dim fileFirma As String = P(0) & "\" & Squadra & "\Scheletri\mail_associato.txt"
														'If Not ControllaEsistenzaFile(fileFirma) Then
														'	fileFirma = Server.MapPath(".") & "\Scheletri\mail_associato.txt"
														'End If
														'Body = gf.LeggeFileIntero(fileFirma)

														'Dim link As String = ""
														'link &= "<a href= """ & Percorso & "?firma=true&codSquadra=" & Squadra & "&id=" & idGiocatore & "&squadra=" & NomeSquadra.Replace(" ", "_") & "&anno=" & Anno & "&genitore=" & Genitore & "&privacy=" & Privacy & "&tipoUtente=1"">"
														'link &= "Click per firmare"
														'link &= "</a>"

														'Body = Body.Replace("***Nominativo Padre***", nomeGenitore)
														'Body = Body.Replace("***Nome societa menu settaggi***", NomeSquadra)
														'Body = Body.Replace("***nome societ&agrave; menu settaggi***", NomeSquadra)
														'Body = Body.Replace("***anno menu settaggi***", Descrizione)
														'Body = Body.Replace("***NOME_LINK_ASSOCIATO****", link)

														''Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, Body, EMail, {})
													Else
														Dim m As New mail
														Dim Oggetto As String = "Richiesta Firma inCalcio"
														'Dim Body As String = gT.EsegueFirma(codSquadra, NomeSquadra, idGiocatore, Anno, Genitore, Privacy)
														'Dim fileFirma As String = P(0) & "\" & Squadra & "\Scheletri\base_firma.txt"
														'If Not ControllaEsistenzaFile(fileFirma) Then
														'	fileFirma = Server.MapPath(".") & "\Scheletri\base_firma.txt"
														'End If
														'Dim Body As String = gf.LeggeFileIntero(fileFirma)

														'Body = Body.Replace("***Nominativo Padre***", nomeGenitore)
														'Body = Body.Replace("***Nome societa menu settaggi***", NomeSquadra)
														'Body = Body.Replace("***anno menu settaggi***", Descrizione)
														'Body = Body.Replace("***cognome menu anagrafica3***", Nominativo)
														'Body = Body.Replace("***Nome menu anagrafica3***", "")

														'Dim link As String = ""
														'link &= "<a href= """ & Percorso & "?firma=true&codSquadra=" & Squadra & "&id=" & idGiocatore & "&squadra=" & NomeSquadra.Replace(" ", "_") & "&anno=" & Anno & "&genitore=" & Genitore & "&privacy=" & Privacy & "&tipoUtente=1"">"
														'link &= "Click per firmare"
														'link &= "</a>"

														'Body = Body.Replace("***NOME_LINK_MAIL****", link)
														'Body = Body.Replace("***nome societ&agrave; menu settaggi***", NomeSquadra)
														'End If

														'' Dim ChiScrive As String = "notifiche@incalcio.cloud"
														'Dim fileDaCopiare As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".html"
														'Dim fileDaCopiarePDF As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".pdf"
														'Dim fileLog As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".log"
														''Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
														'gf.CreaDirectoryDaPercorso(fileDaCopiare)

														'Dim fileDaCopiarePrivacy As String = P(0) & "\" & Squadra & "\Firme\privacy_" & Anno & "_" & idGiocatore & ".html"
														'Dim fileDaCopiarePrivacyPDF As String = P(0) & "\" & Squadra & "\Firme\privacy_" & Anno & "_" & idGiocatore & ".pdf"

														'Dim fileScheletro As String = P(0) & Squadra & "\Scheletri\base_iscrizione_.txt"
														'If Not ControllaEsistenzaFile(fileScheletro) Then
														'	fileScheletro = Server.MapPath(".") & "\Scheletri\base_iscrizione_.txt"
														'End If

														'Dim fileScheletroPrivacy As String = P(0) & Squadra & "\Scheletri\base_privacy.txt"
														'If Not ControllaEsistenzaFile(fileScheletroPrivacy) Then
														'	fileScheletroPrivacy = Server.MapPath(".") & "\Scheletri\base_privacy.txt"
														'End If

														' If ControllaEsistenzaFile(fileScheletro) And ControllaEsistenzaFile(fileScheletroPrivacy) Then

														ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Esegue file firme: " & Squadra & " - Id Giocatore: " & idGiocatore & " - Anno: " & Anno)
														Dim fileFirme As String = gT1.EsegueFileFirme(Server.MapPath("."), Squadra, idGiocatore, Anno, True)
														ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Eseguito file firme")

														'Try
														'Dim fileFirme As String = gf.LeggeFileIntero(fileScheletro)
														'fileFirme = RiempieFileFirme(fileFirme, Anno, idGiocatore, Rec, Conn, Connessione, NomeSquadra, P, Descrizione)
														'If Maggiorenne = "S" Then
														'	fileFirme = fileFirme.Replace("---VIS_PADRE---", "none")

														'	fileFirme = fileFirme.Replace("---VIS_MADRE---", "none")
														'Else
														'	If GenitoriSeparati = "S" Then
														'		If AffidamentoCongiunto = "S" Then
														'			If iscrFirmaEntrambi = "S" Then
														'				fileFirme = fileFirme.Replace("---VIS_PADRE---", "block")

														'				fileFirme = fileFirme.Replace("---VIS_MADRE---", "block")
														'			Else
														'				If ceGenitore1 = "S" Then
														'					fileFirme = fileFirme.Replace("---VIS_PADRE---", "block")

														'					fileFirme = fileFirme.Replace("---VIS_MADRE---", "none")
														'				Else
														'					fileFirme = fileFirme.Replace("---VIS_PADRE---", "none")

														'					fileFirme = fileFirme.Replace("---VIS_MADRE---", "block")
														'				End If
														'			End If
														'		Else
														'			If idTutore = "1" Then
														'				fileFirme = fileFirme.Replace("---VIS_MADRE---", "none")
														'			Else
														'				fileFirme = fileFirme.Replace("---VIS_PADRE---", "none")
														'			End If
														'		End If
														'	Else
														'		If iscrFirmaEntrambi = "S" Then
														'			fileFirme = fileFirme.Replace("---VIS_PADRE---", "block")

														'			fileFirme = fileFirme.Replace("---VIS_MADRE---", "block")
														'		Else
														'			If ceGenitore1 = "S" Then
														'				fileFirme = fileFirme.Replace("---VIS_PADRE---", "block")

														'				fileFirme = fileFirme.Replace("---VIS_MADRE---", "none")
														'			Else
														'				fileFirme = fileFirme.Replace("---VIS_PADRE---", "none")

														'				fileFirme = fileFirme.Replace("---VIS_MADRE---", "block")
														'			End If
														'		End If
														'	End If
														'End If

														'gf.EliminaFileFisico(fileDaCopiare)
														'gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
														'gf.ScriveTestoSuFileAperto(fileFirme)
														'gf.ChiudeFileDiTestoDopoScrittura()

														'Dim filePrivacy As String = gf.LeggeFileIntero(fileScheletroPrivacy)
														'filePrivacy = RiempieFilePrivacy(filePrivacy, Anno, idGiocatore, Rec, Conn, Connessione, NomeSquadra, P, Descrizione)
														gT1 = Nothing

														Dim gT2 As New GestioneTags(Server.MapPath("."))
														ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Esegue file privacy: " & Squadra & " - Id Giocatore: " & idGiocatore & " - Anno: " & Anno)
														Dim filePrivacy As String = gT2.EsegueFilePrivacy(Server.MapPath("."), Squadra, idGiocatore, Anno, True)
														ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Eseguito file privacy")
														gT2 = Nothing

														'gf.EliminaFileFisico(fileDaCopiarePrivacy)
														'gf.ApreFileDiTestoPerScrittura(fileDaCopiarePrivacy)
														'gf.ScriveTestoSuFileAperto(filePrivacy)
														'gf.ChiudeFileDiTestoDopoScrittura()

														'File.Copy(fileDaCopiare, fileDaCopiare2)
														'Dim pp As New pdfGest
														'Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)
														'Ritorno = pp.ConverteHTMLInPDF(fileDaCopiarePrivacy, fileDaCopiarePrivacyPDF, fileLog)

														'If Ritorno = "*" Then
														Dim gT3 As New GestioneTags(Server.MapPath("."))
														Dim fileDaCopiarePDF As String = Server.MapPath(".") & "\Appoggio\iscrizione_" & Anno & "_" & idGiocatore & ".pdf"
														Dim fileDaCopiarePrivacyPDF As String = Server.MapPath(".") & "\Appoggio\Privacy_" & Anno & "_" & idGiocatore & ".pdf"

														Dim filesDaAllegare() As String = {fileDaCopiarePDF, fileDaCopiarePrivacyPDF}
														''gf.EliminaFileFisico(fileDaCopiare)
														ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "Invio mail")
														' Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, gT3.EsegueMailAssociato(Server.MapPath("."), Squadra, idGiocatore, Anno, QualeGenitore, Privacy), EMail, filesDaAllegare)
														Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, gT3.EsegueMailAssociato(Server.MapPath("."), Squadra, idGiocatore, Anno, QualeGenitore, QualeFirma), EMail, filesDaAllegare)
														gT3 = Nothing
														'End If

														gf = Nothing

														If Ritorno = "" Then Ritorno = "*"
														'Catch ex As Exception
														'	Ritorno = StringaErrore & " " & ex.Message
														'End Try
														'Else
														'	Ritorno = StringaErrore & " Scheletro iscrizione oppure privacy non trovato"
														'	End If
														'gf = Nothing
														'Ritorno = "*"
													End If
												End If
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

				'If Ritorno = "*" Then
				'	Sql = "commit"
				'	Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				'Else
				'	Sql = "rollback"
				'	Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				'End If
			End If
		End If
		ScriveLog(Server.MapPath("."), Squadra, "RichiedeFirma", "--------------------------------")

		Return Ritorno
	End Function

	'Private Function RiempieFilePrivacy(Contenuto As String, Anno As String, idGiocatore As String, Rec As Object, Conn As Object, Connessione As String, Squadra As String, p() As String, DescAnno As String) As String
	'	Dim c As New CriptaFiles
	'	p(2) = p(2).Replace(vbCrLf, "")
	'	If (Strings.Right(p(2), 1) <> "/") Then
	'		p(2) = p(2) & "/"
	'	End If

	'	Dim Sql As String = "Select * From Anni Where idAnno=" & Anno
	'	Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
	'	If TypeOf (Rec) Is String Then
	'		Contenuto = Rec
	'	Else
	'		If Not Rec.Eof() Then
	'			Dim NomePolisportiva As String = "" & Rec("NomePolisportiva").value
	'			Dim Mail As String = "" & Rec("Mail").value
	'			Dim Telefono As String = "" & Rec("Telefono").value
	'			Dim Indirizzo As String = "" & Rec("Indirizzo").Value
	'			Dim CodiceFiscale As String = "" & Rec("CodiceFiscale").Value

	'			Contenuto = Contenuto.Replace("***Nome Societ&agrave;***", NomePolisportiva)
	'			Contenuto = Contenuto.Replace("***indirizzo***", Indirizzo)
	'			Contenuto = Contenuto.Replace("***Mail***", Mail)
	'			Contenuto = Contenuto.Replace("***Cofice Fiscale***", CodiceFiscale)
	'		Else
	'			Contenuto = Contenuto.Replace("***Nome Societ&agrave;***", "")
	'			Contenuto = Contenuto.Replace("***indirizzo***", "")
	'			Contenuto = Contenuto.Replace("***Mail***", "")
	'			Contenuto = Contenuto.Replace("***Cofice Fiscale***", "")
	'		End If
	'	End If

	'	Return Contenuto
	'End Function

	'Private Function RiempieFileFirme(Contenuto As String, Anno As String, idGiocatore As String, Rec As Object, Conn As Object, Connessione As String, Squadra As String, p() As String, DescAnno As String) As String
	'	Dim c As New CriptaFiles

	'	'Dim gf As New GestioneFilesDirectory
	'	'Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
	'	'Dim p() As String = paths.Split(";")
	'	p(2) = p(2).Replace(vbCrLf, "")
	'	If (Strings.Right(p(2), 1) <> "/") Then
	'		p(2) = p(2) & "/"
	'	End If

	'	Dim Sql As String = "Select * From Anni Where idAnno=" & Anno
	'	Dim NomeSquadra As String = ""

	'	Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
	'	If TypeOf (Rec) Is String Then
	'		Contenuto = Rec
	'	Else
	'		If Not Rec.Eof() Then
	'			NomeSquadra = "" & Rec("NomeSquadra").Value
	'			Dim NomePolisportiva As String = "" & Rec("NomePolisportiva").value
	'			Dim NomeCampo As String = "" & Rec("CampoSquadra").value
	'			Dim Mail As String = "" & Rec("Mail").value
	'			Dim Telefono As String = "" & Rec("Telefono").value
	'			Dim SitoWeb As String = "" & Rec("SitoWeb").value
	'			Dim Indirizzo As String = "" & Rec("Indirizzo").Value
	'			Dim CodiceFiscale As String = "" & Rec("CodiceFiscale").Value
	'			Dim PIva As String = "" & Rec("PIva").Value

	'			Contenuto = Contenuto.Replace("***Anno menu settaggi***", DescAnno)
	'			Contenuto = Contenuto.Replace("***nome societa menu settaggi***", NomePolisportiva)
	'			Contenuto = Contenuto.Replace("***nome Campo menu settaggi***", NomeCampo)
	'			Contenuto = Contenuto.Replace("***Telefono - mail - sito web menu settaggi***", Mail & ", " & Telefono & ", " & SitoWeb)
	'			Contenuto = Contenuto.Replace("***indirizzo menu settaggi tab Dati Generali***", Indirizzo)
	'			Contenuto = Contenuto.Replace("***codice fiscale menu settaggi***", CodiceFiscale)
	'			Contenuto = Contenuto.Replace("***partita iva menu settaggi***", PIva)
	'		Else
	'			Contenuto = Contenuto.Replace("***Anno menu settaggi***", Anno)
	'			Contenuto = Contenuto.Replace("***nome societa menu settaggi***", "")
	'			Contenuto = Contenuto.Replace("***nome Campo menu settaggi***", "")
	'			Contenuto = Contenuto.Replace("***Telefono - mail - sito web menu settaggi***", "")
	'			Contenuto = Contenuto.Replace("***indirizzo menu settaggi tab Dati Generali***", "")
	'			Contenuto = Contenuto.Replace("***codice fiscale menu settaggi***", "")
	'			Contenuto = Contenuto.Replace("***partita iva menu settaggi***", "")
	'		End If

	'		Sql = "Select * From Giocatori Where idGiocatore=" & idGiocatore
	'		Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
	'		If TypeOf (Rec) Is String Then
	'			Contenuto = Rec
	'		Else
	'			If Not Rec.Eof() Then
	'				Dim Cognome As String = "" & Rec("Cognome").value
	'				Dim Nome As String = "" & Rec("Nome").value
	'				Dim ddn As String = "" & Rec("DataDiNascita").value
	'				If ddn <> "" Then
	'					Dim d() As String = ddn.Split("-")
	'					ddn = d(2) & "/" & d(1) & "/" & d(0)
	'				End If
	'				Dim DataDiNascita As String = ddn
	'				Dim CodFisc As String = "" & Rec("CodFiscale").value
	'				Dim Maschio As String = "" & Rec("Maschio").value
	'				Dim Indirizzo As String = "" & Rec("Indirizzo").value
	'				Dim Citta As String = "" & Rec("Citta").value
	'				Dim EMail As String = "" & Rec("EMail").value
	'				Dim TelefonoGioc As String = "" & Rec("Telefono").value
	'				Dim Cap As String = "" & Rec("Cap").value
	'				Dim CittaNascita As String = "" & Rec("CittaNascita").value

	'				If Maschio = "M" Then
	'					Maschio = "Maschile"
	'				Else
	'					Maschio = "Femminile"
	'				End If

	'				Contenuto = Contenuto.Replace("****cognome menu anagrafica3***", Cognome)
	'				Contenuto = Contenuto.Replace("***Nome menu anagrafica3***", Nome)
	'				Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica3***", DataDiNascita)
	'				Contenuto = Contenuto.Replace("***non c'&egrave;***", "")
	'				Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica3***", CodFisc)
	'				Contenuto = Contenuto.Replace("***sesso menu anagrafica***", Maschio)
	'				Contenuto = Contenuto.Replace("****indirizzo menu anagrafica3***", Indirizzo)
	'				Contenuto = Contenuto.Replace("***citta3***", Citta)
	'				Contenuto = Contenuto.Replace("***?***", "")
	'				Contenuto = Contenuto.Replace("*** mail menu anagrafica3***", EMail)
	'				Contenuto = Contenuto.Replace("***telefono menu anagrafica3***", TelefonoGioc)
	'				Contenuto = Contenuto.Replace("***?Cap3***", Cap)
	'				Contenuto = Contenuto.Replace("***Citta di nascita3***", CittaNascita)
	'			Else
	'				Contenuto = Contenuto.Replace("****cognome menu anagrafica3***", "")
	'				Contenuto = Contenuto.Replace("***Nome menu anagrafica3***", "")
	'				Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica3***", "")
	'				Contenuto = Contenuto.Replace("***non c'&egrave;***", "")
	'				Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica3***", "")
	'				Contenuto = Contenuto.Replace("***sesso menu anagrafica***", "")
	'				Contenuto = Contenuto.Replace("****indirizzo menu anagrafica3***", "")
	'				Contenuto = Contenuto.Replace("***citta3***", "")
	'				Contenuto = Contenuto.Replace("***?***", "")
	'				Contenuto = Contenuto.Replace("*** mail menu anagrafica3***", "")
	'				Contenuto = Contenuto.Replace("***telefono menu anagrafica3***", "")
	'				Contenuto = Contenuto.Replace("***?Cap3***", "")
	'				Contenuto = Contenuto.Replace("***Citta di nascita3***", "")
	'			End If
	'		End If

	'		Sql = "Select * From GiocatoriDettaglio Where idGiocatore=" & idGiocatore
	'		Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
	'		If TypeOf (Rec) Is String Then
	'			Contenuto = Rec
	'		Else
	'			If Not Rec.Eof() Then
	'				Dim Genitore1 As String = "" & Rec("Genitore1").value
	'				Dim Mail1 As String = "" & Rec("MailGenitore1").value
	'				Dim Telefono1 As String = "" & Rec("TelefonoGenitore1").value
	'				Dim Gen1() As String = Genitore1.Split(" ")
	'				If Gen1.Length = 1 Then
	'					ReDim Preserve Gen1(2)
	'				End If
	'				Dim ddn As String = "" & Rec("DataDiNascita1").Value
	'				If ddn <> "" Then
	'					Dim d() As String = ddn.Split("-")
	'					ddn = d(2) & "/" & d(1) & "/" & d(0)
	'				End If
	'				Dim DataDiNascita1 As String = ddn
	'				Dim CittaNascita1 As String = "" & Rec("CittaNascita1").Value
	'				Dim CodFiscale1 As String = "" & Rec("CodFiscale1").Value
	'				Dim Citta1 As String = "" & Rec("Citta1").Value
	'				Dim Cap1 As String = "" & Rec("Cap1").Value
	'				Dim Indirizzo1 As String = "" & Rec("Indirizzo1").Value

	'				Dim Genitore2 As String = "" & Rec("Genitore2").value
	'				Dim Mail2 As String = "" & Rec("MailGenitore2").value
	'				Dim Telefono2 As String = "" & Rec("TelefonoGenitore2").value
	'				Dim Gen2() As String = Genitore2.Split(" ")
	'				If Gen2.Length = 1 Then
	'					ReDim Preserve Gen2(2)
	'				End If
	'				ddn = "" & Rec("DataDiNascita2").Value
	'				If ddn <> "" Then
	'					Dim d() As String = ddn.Split("-")
	'					ddn = d(2) & "/" & d(1) & "/" & d(0)
	'				End If
	'				Dim DataDiNascita2 As String = ddn
	'				Dim CittaNascita2 As String = "" & Rec("CittaNascita2").Value
	'				Dim CodFiscale2 As String = "" & Rec("CodFiscale2").Value
	'				Dim Citta2 As String = "" & Rec("Citta2").Value
	'				Dim Cap2 As String = "" & Rec("Cap2").Value
	'				Dim Indirizzo2 As String = "" & Rec("Indirizzo2").Value

	'				Contenuto = Contenuto.Replace("****cognome menu anagrafica1***", Gen1(1))
	'				Contenuto = Contenuto.Replace("***Nome menu anagrafica1***", Gen1(0))
	'				Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica1***", DataDiNascita1)
	'				Contenuto = Contenuto.Replace("***Citta di nascita1***", CittaNascita1)
	'				Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica1***", CodFiscale1)
	'				Contenuto = Contenuto.Replace("****indirizzo menu anagrafica1***", Indirizzo1)
	'				Contenuto = Contenuto.Replace("***citta1***", Citta1)
	'				Contenuto = Contenuto.Replace("***Cap1***", Cap1)
	'				Contenuto = Contenuto.Replace("*** mail menu anagrafica1***", Mail1)
	'				Contenuto = Contenuto.Replace("***telefono menu anagrafica1***", Indirizzo1)

	'				Contenuto = Contenuto.Replace("****cognome menu anagrafica2***", Gen2(1))
	'				Contenuto = Contenuto.Replace("***Nome menu anagrafica2***", Gen2(0))
	'				Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica2***", DataDiNascita2)
	'				Contenuto = Contenuto.Replace("***Citta di nascita2***", CittaNascita2)
	'				Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica2***", CodFiscale2)
	'				Contenuto = Contenuto.Replace("****indirizzo menu anagrafica2***", Indirizzo2)
	'				Contenuto = Contenuto.Replace("***citta2***", Citta2)
	'				Contenuto = Contenuto.Replace("***Cap2***", Cap2)
	'				Contenuto = Contenuto.Replace("*** mail menu anagrafica2***", Mail2)
	'				Contenuto = Contenuto.Replace("***telefono menu anagrafica2***", Indirizzo2)
	'			Else
	'				Contenuto = Contenuto.Replace("****cognome menu anagrafica1***", "")
	'				Contenuto = Contenuto.Replace("***Nome menu anagrafica1***", "")
	'				Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica1***", "")
	'				Contenuto = Contenuto.Replace("***Citta Nascita 1***", "")
	'				Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica1***", "")
	'				Contenuto = Contenuto.Replace("****indirizzo menu anagrafica1***", "")
	'				Contenuto = Contenuto.Replace("***citta1***", "")
	'				Contenuto = Contenuto.Replace("***Cap1***", "")
	'				Contenuto = Contenuto.Replace("*** mail menu anagrafica1***", "")
	'				Contenuto = Contenuto.Replace("***telefono menu anagrafica1***", "")

	'				Contenuto = Contenuto.Replace("****cognome menu anagrafica2***", "")
	'				Contenuto = Contenuto.Replace("***Nome menu anagrafica2***", "")
	'				Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica2***", "")
	'				Contenuto = Contenuto.Replace("***Citta di nascita2***", "")
	'				Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica2***", "")
	'				Contenuto = Contenuto.Replace("****indirizzo menu anagrafica2***", "")
	'				Contenuto = Contenuto.Replace("***citta2***", "")
	'				Contenuto = Contenuto.Replace("***Cap2***", "")
	'				Contenuto = Contenuto.Replace("*** mail menu anagrafica2***", "")
	'				Contenuto = Contenuto.Replace("***telefono menu anagrafica2***", "")
	'			End If
	'		End If

	'		'Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore
	'		'Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
	'		'If TypeOf (Rec) Is String Then
	'		'	Contenuto = Rec
	'		'Else
	'		'	Do Until Rec.Eof()
	'		'		Select Case Rec("idGenitore").value
	'		'			Case 1
	'		'				Contenuto = Contenuto.Replace("***data firma1***", Rec("DataFirma").value)
	'		'			Case 2
	'		'				Contenuto = Contenuto.Replace("***data firma2***", Rec("DataFirma").value)
	'		'			Case 3
	'		'				Contenuto = Contenuto.Replace("***data firma3***", Rec("DataFirma").value)
	'		'		End Select

	'		'		Rec.movenext
	'		'	Loop
	'		'	Contenuto = Contenuto.Replace("***data firma1***", "")
	'		'	Contenuto = Contenuto.Replace("***data firma2***", "")
	'		'	Contenuto = Contenuto.Replace("***data firma3***", "")
	'		'End If

	'		Dim Datella As String = "* " & Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

	'		Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=1"
	'		Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
	'		If TypeOf (Rec) Is String Then
	'			Contenuto = Rec
	'		Else
	'			If Not Rec.Eof() Then
	'				Datella = "" & Rec("DataFirma").Value
	'				If Datella.Contains(" ") Then
	'					Datella = Mid(Datella, 1, Datella.IndexOf(" "))
	'				End If
	'			End If
	'			Rec.Close()
	'		End If
	'		Contenuto = Contenuto.Replace("***data firma2***", Datella)

	'		Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=2"
	'		Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
	'		If TypeOf (Rec) Is String Then
	'			Contenuto = Rec
	'		Else
	'			If Not Rec.Eof() Then
	'				Datella = "" & Rec("DataFirma").Value
	'				If Datella.Contains(" ") Then
	'					Datella = Mid(Datella, 1, Datella.IndexOf(" "))
	'				End If
	'			End If
	'			Rec.Close()
	'		End If
	'		Contenuto = Contenuto.Replace("***data firma3***", Datella)

	'		Dim gf As New GestioneFilesDirectory
	'		Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
	'		pp = pp.Trim()
	'		If Strings.Right(pp, 1) = "\" Then
	'			pp = Mid(pp, 1, pp.Length - 1)
	'		End If
	'		Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

	'		Dim pathFirma1 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_1.kgb"
	'		Dim urlFirma1 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_1.kgb"
	'		Dim pathFirmaConv1 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_1_" & Esten & ".png"
	'		Dim urlFirmaConv1 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_1_" & Esten & ".png"

	'		Dim pathFirma2 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_2.kgb"
	'		Dim urlFirma2 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_2.kgb"
	'		Dim pathFirmaConv2 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_2_" & Esten & ".png"
	'		Dim urlFirmaConv2 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_2_" & Esten & ".png"

	'		Dim pathFirma3 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_3.kgb"
	'		Dim urlFirma3 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_3.kgb"
	'		Dim pathFirmaConv3 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_3_" & Esten & ".png"
	'		Dim urlFirmaConv3 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_3_" & Esten & ".png"

	'		Dim nomeImm As String = p(2) & Squadra.Replace(" ", "_") & "/Societa/" & Anno & "_1.kgb"
	'		Dim pathImm As String = pp & "\" & Squadra.Replace(" ", "_") & "\Societa\" & Anno & "_1.kgb"
	'		If ControllaEsistenzaFile(pathImm) Then
	'			Dim nomeImmConv As String = p(2) & "/" & NomeSquadra.Replace(" ", "_") & "/Societa/Societa_1.png"
	'			Dim pathImmConv As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\Societa_1.png"
	'			c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)

	'			Contenuto = Contenuto.Replace("***immagine logo menu settaggi***", "<img src=""" & nomeImmConv & """ style=""width: 100px; height: 100px;"" />")
	'		Else
	'			Contenuto = Contenuto.Replace("***immagine logo menu settaggi***", "")
	'		End If

	'		nomeImm = p(2) & Squadra.Replace(" ", "_") & "/Societa/" & Anno & "_2.kgb"
	'		pathImm = pp & "\" & Squadra.Replace(" ", "_") & "\Societa\" & Anno & "_2.kgb"
	'		If ControllaEsistenzaFile(pathImm) Then
	'			Dim nomeImmConv As String = p(2) & "/" & NomeSquadra.Replace(" ", "_") & "/Societa/Societa_2.png"
	'			Dim pathImmConv As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\Societa_2.png"
	'			c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)

	'			Contenuto = Contenuto.Replace("***immagine logo affiliazione menu settaggi***", "<img src=""" & nomeImmConv & """ style=""width: 100px; height: 100px;"" />")
	'		Else
	'			Contenuto = Contenuto.Replace("***immagine logo affiliazione menu settaggi***", "")
	'		End If

	'		If ControllaEsistenzaFile(urlFirma1) Then
	'			c.DecryptFile(CryptPasswordString, urlFirma1, urlFirmaConv1)
	'			Contenuto = Contenuto.Replace("***firma padre***", "FIRMA: <img src=""" & pathFirmaConv1 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
	'		Else
	'			Contenuto = Contenuto.Replace("***firma padre***", "FIRMA: " & "")
	'		End If
	'		If ControllaEsistenzaFile(urlFirma2) Then
	'			c.DecryptFile(CryptPasswordString, urlFirma2, urlFirmaConv2)
	'			Contenuto = Contenuto.Replace("***firma madre***", "FIRMA: <img src=""" & pathFirmaConv2 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
	'		Else
	'			Contenuto = Contenuto.Replace("***firma madre***", "FIRMA: " & "")
	'		End If
	'		If ControllaEsistenzaFile(urlFirma3) Then
	'			c.DecryptFile(CryptPasswordString, urlFirma3, urlFirmaConv3)
	'			Contenuto = Contenuto.Replace("***firma giocatore***", "FIRMA: <img src=""" & pathFirmaConv3 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
	'		Else
	'			Contenuto = Contenuto.Replace("***firma giocatore***", "FIRMA: " & "")
	'		End If

	'		Dim pathFirmaPrivacy1 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_1_P.kgb"
	'		Dim urlFirmaPrivacy1 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_1_P.kgb"
	'		Dim pathFirmaConvPrivacy1 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_1_" & Esten & "_P.png"
	'		Dim urlFirmaConvPrivacy1 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_1_" & Esten & "_P.png"
	'		If ControllaEsistenzaFile(urlFirmaPrivacy1) Then
	'			c.DecryptFile(CryptPasswordString, urlFirmaPrivacy1, urlFirmaConvPrivacy1)
	'			Contenuto = Contenuto.Replace("***firma privacy padre***", "FIRMA: <img src=""" & pathFirmaConvPrivacy1 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
	'		Else
	'			Contenuto = Contenuto.Replace("***firma privacy padre***", "FIRMA: " & "")
	'		End If

	'		Dim pathFirmaPrivacy2 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_2_P.kgb"
	'		Dim urlFirmaPrivacy2 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_2_P.kgb"
	'		Dim pathFirmaConvPrivacy2 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_2_" & Esten & "_P.png"
	'		Dim urlFirmaConvPrivacy2 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_2_" & Esten & "_P.png"
	'		If ControllaEsistenzaFile(urlFirmaPrivacy2) Then
	'			c.DecryptFile(CryptPasswordString, urlFirmaPrivacy2, urlFirmaConvPrivacy2)
	'			Contenuto = Contenuto.Replace("***firma privacy madre***", "FIRMA: <img src=""" & pathFirmaConvPrivacy2 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
	'		Else
	'			Contenuto = Contenuto.Replace("***firma privacy madre***", "FIRMA: " & "")
	'		End If

	'	End If

	'	'Contenuto &= "<hr />Stampato tramite InCalcio, software per la gestione delle società di calcio - www.incalcio.it - info@incalcio.it"

	'	Return Contenuto
	'End Function

	<WebMethod()>
	Public Function SalvaGiocatoriNote(Squadra As String, ByVal idGiocatore As String, Notelle As String) As String
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

				Sql = "Delete From GiocatoriNote Where idGiocatore=" & idGiocatore
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Ritorno = "*" Then
					Sql = "Insert Into GiocatoriNote Values (" & idGiocatore & ", '" & Notelle & "')"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaGiocatoriNote(Squadra As String, ByVal idGiocatore As String) As String
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

				Sql = "Select * From GiocatoriNote Where idGiocatore=" & idGiocatore
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = ""
					Else
						Ritorno = Rec("Nota").Value
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaGiocatoriCategoria(Squadra As String, ByVal idAnno As String, ByVal idCategoria As String) As String
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
				Dim Altro As String = ""

				If idCategoria <> "-1" Then
					Altro = IIf(TipoDB = "SQLSERVER", "And CharIndex('" & idCategoria & "-', Categorie) > 0", "And Instr(Categorie, '" & idCategoria & "-') > 0")
				End If

				Try
					Sql = "SELECT Giocatori.idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, idCategoria2, Categorie.Descrizione As Categoria2, idCategoria3, Cat3.Descrizione As Categoria3, Cat1.Descrizione As Categoria1, " &
						"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne, " &
						"  " & IIf(TipoDB = "SQLSERVER", "IsNull(Cat4.ScadenzaCertificatoMedico,'')", "Coalesce(Cat4.ScadenzaCertificatoMedico,'')") & " As ScadenzaCertificatoMedico, Cat4.CertificatoMedico, CodiceTessera " &
						"FROM Giocatori " &
						"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria2 And Categorie.idAnno=Giocatori.idAnno " &
						"Left Join Categorie As Cat3 On Cat3.idCategoria=Giocatori.idCategoria3 And Cat3.idAnno=Giocatori.idAnno " &
						"Left Join GiocatoriDettaglio As Cat4 On Cat4.idGiocatore=Giocatori.idGiocatore " &
						"Left Join Categorie As Cat1 On Cat1.idCategoria=Giocatori.idCategoria And Cat1.idAnno=Giocatori.idAnno " &
						"Left Join [Generale].[dbo].[GiocatoriTessereNFC] As NFC On NFC.idGiocatore=Giocatori.idGiocatore And NFC.CodSquadra = '" & Squadra & "' " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " " & Altro & " " &
						"And RapportoCompleto = 'S' " &
						"Order By Cognome, Nome"
					' "Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And (Giocatori.idCategoria=" & idCategoria & " Or Giocatori.idCategoria2=" & idCategoria & " Or Giocatori.idCategoria3=" & idCategoria & ") " &
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Dim dat As Date = Nothing
								Dim Scaduto As String = "S"

								If Rec("ScadenzaCertificatoMedico").Value <> "" Then
									dat = Convert.ToDateTime(Rec("ScadenzaCertificatoMedico").Value)
									Dim days As Long = DateDiff(DateInterval.Day, dat, Now)
									If days < 0 Then
										Scaduto = "N"
									End If
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
									Rec("Cap").Value.ToString & ";" &
									Rec("CittaNascita").Value.ToString & ";" &
									Rec("Maggiorenne").Value & ";" &
									Rec("CertificatoMedico").Value & ";" &
									Scaduto & ";" &
									Rec("CodiceTessera").Value & ";" &
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					Sql = "SELECT idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, -1 As idCategoria2, '' As Categoria2, -1 As idCategoria3, '' As Categoria3, Categorie.Descrizione As Categoria1, " &
						"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne, CodiceTesseraNFC " &
						"FROM Giocatori " &
						"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
						"Left Join [Generale].[dbo].[GiocatoriTessereNFC] As NFC On NFC.idGiocatore=Giocatori.idGiocatore And NFC.CodSquadra = '" & Squadra & "' " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And Giocatori.idCategoria=" & idCategoria & " " &
						"And RapportoCompleto = 'S' " &
						"Order By Ruoli.idRuolo, Cognome, Nome"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
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
									Rec("CodiceTessera").Value.ToString & ";" &
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					Sql = "SELECT Giocatori.idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, -1 As idCategoria2, '' As Categoria2, -1 As idCategoria3, '' As Categoria3, Categorie.Descrizione As Categoria1, " &
						"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne, CodiceTessera " &
						"FROM Giocatori " &
						"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
						"Left Join [Generale].[dbo].[GiocatoriTessereNFC] As NFC On NFC.idGiocatore=Giocatori.idGiocatore  And NFC.CodSquadra='" & Squadra & "' " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And " & IIf(TipoDB = "SQLSERVER", "CharIndex('" & idCategoria & "-', Categorie) = 0 ", "Instr(Categorie, '" & idCategoria & "-') = 0 ") & " " &
						"And Giocatori.RapportoCompleto = 'S' " &
						"Order By Ruoli.idRuolo, Cognome, Nome"
					' "Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And Giocatori.idCategoria<>" & idCategoria & " " &
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
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
									Rec("CodiceTessera").Value.ToString & ";" &
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Rec2 As Object
				Dim Sql As String = ""

				If Ritorno = "" Then
					Dim Ok As Boolean = True
					Dim gf As New GestioneFilesDirectory
					Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "\" Then
						Percorso &= "\"
					End If

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

					Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
					Dim P() As String = paths.Split(";")
					If Strings.Right(P(0), 1) <> "\" Then
						P(0) &= "\"
					End If
					Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
					If Strings.Right(P(2), 1) <> "/" Then
						P(2) &= "/"
					End If
					Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
					Dim PathBaseMultimedia As String = P(3) ' pathMultimedia.Replace("Allegati", "Multimedia")
					Dim pathHTTPMultimedia As String = P(2)
					gf = Nothing

					If Ok Then
						Try
							Sql = "SELECT Giocatori.idGiocatore, Ruoli.idRuolo As idR, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione, Giocatori.EMail, Giocatori.Telefono, Giocatori.Soprannome, Giocatori.DataDiNascita, Giocatori.Indirizzo, " &
								"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, Giocatori.idCategoria2 As idCategoria2, Categorie2.Descrizione As Categoria2, " &
								"Giocatori.idCategoria3 As idCategoria3, Categorie3.Descrizione As Categoria3, Categorie.Descrizione As Categoria1, Giocatori.Categorie, " &
								"Giocatori.RapportoCompleto, Giocatori.idTaglia, Min(KitGiocatori.idTipoKit) As idTipologiaKit, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne, " &
								" " & IIf(TipoDB = "SQLSERVER", "IsNull(GiocatoriSemafori.Semaforo1,'')", "Coalesce(GiocatoriSemafori.Semaforo1,'')") & " As Semaforo1, GiocatoriSemafori.Titolo1, " &
								" " & IIf(TipoDB = "SQLSERVER", "IsNull(GiocatoriSemafori.Semaforo2,'')", "Coalesce(GiocatoriSemafori.Semaforo2,'')") & " As Semaforo2, GiocatoriSemafori.Titolo2, " &
								" " & IIf(TipoDB = "SQLSERVER", "IsNull(GiocatoriSemafori.Smeaforo3,'')", "Coalesce(GiocatoriSemafori.Smeaforo3,'')") & " As Semaforo3, GiocatoriSemafori.Titolo3, " &
								" " & IIf(TipoDB = "SQLSERVER", "IsNull(GiocatoriSemafori.Semaforo4,'')", "Coalesce(GiocatoriSemafori.Semaforo4,'')") & " As Semaforo4, GiocatoriSemafori.Titolo4, " &
								" " & IIf(TipoDB = "SQLSERVER", "IsNull(GiocatoriSemafori.Semaforo5,'')", "Coalesce(GiocatoriSemafori.Semaforo5,'')") & " As Semaforo5, GiocatoriSemafori.Titolo5, CodiceTessera, " &
								" " & IIf(TipoDB = "SQLSERVER", "IsNull(GiocatoriDettaglio.MailGenitore1,'')", "Coalesce(GiocatoriDettaglio.MailGenitore1,'')") & " As MailGenitore1, " &
								" " & IIf(TipoDB = "SQLSERVER", "IsNull(GiocatoriDettaglio.MailGenitore2,'')", "Coalesce(GiocatoriDettaglio.MailGenitore2,'')") & " As MailGenitore2, " &
								"UtentiPadre.idGiocatore As AmministratiPadre, UtentiMadre.idGiocatore As AmministratiMadre " &
								"FROM Giocatori " &
								"Left Join KitGiocatori On Giocatori.idGiocatore=KitGiocatori.idGiocatore " &
								"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
								"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
								"Left Join Categorie As Categorie2 On Categorie2.idCategoria=Giocatori.idCategoria2 And Categorie2.idAnno=Giocatori.idAnno " &
								"Left Join Categorie As Categorie3 On Categorie3.idCategoria=Giocatori.idCategoria3 And Categorie3.idAnno=Giocatori.idAnno " &
								"Left Join GiocatoriSemafori On Giocatori.idGiocatore = GiocatoriSemafori.idGiocatore " &
								"Left Join [Generale].[dbo].[GiocatoriTessereNFC] As NFC On NFC.idGiocatore=Giocatori.idGiocatore And NFC.CodSquadra = '" & Squadra & "' " &
								"Left Join GiocatoriDettaglio On GiocatoriDettaglio.idGiocatore = Giocatori.idGiocatore " &
								"Left Join [Generale].[dbo].[Utenti] As UtentiPadre On GiocatoriDettaglio.MailGenitore1=UtentiPadre.Utente " &
								"Left Join [Generale].[dbo].[Utenti] As UtentiMadre On GiocatoriDettaglio.MailGenitore2=UtentiMadre.Utente " &
								"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " " &
								"Group By Giocatori.idGiocatore, Ruoli.idRuolo, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione, Giocatori.EMail, Giocatori.Telefono, Giocatori.Soprannome, Giocatori.DataDiNascita, Giocatori.Indirizzo, CodFiscale, Maschio, " &
								"Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, Giocatori.idCategoria2, Categorie2.Descrizione, Giocatori.idCategoria3, Categorie3.Descrizione, Categorie.Descrizione, " &
								"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.idTaglia, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne, " &
								"GiocatoriSemafori.Semaforo1, GiocatoriSemafori.Titolo1, GiocatoriSemafori.Semaforo2, GiocatoriSemafori.Titolo2, GiocatoriSemafori.Smeaforo3, GiocatoriSemafori.Titolo3, " &
								"GiocatoriSemafori.Semaforo4, GiocatoriSemafori.Titolo4, GiocatoriSemafori.Semaforo5, GiocatoriSemafori.Titolo5, CodiceTessera, GiocatoriDettaglio.MailGenitore1, GiocatoriDettaglio.MailGenitore2, " &
								"UtentiPadre.idGiocatore, UtentiMadre.idGiocatore " &
								"Order By Giocatori.Cognome, Giocatori.Nome"
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)

							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
								Else
									Ritorno = ""
									Do Until Rec.Eof()
										Dim Semaforo1 As String = ""
										Dim Semaforo2 As String = ""
										Dim Semaforo3 As String = ""
										Dim Semaforo4 As String = ""
										Dim Semaforo5 As String = ""

										If Rec("Semaforo1").Value = "" Then
											Semaforo1 = "rosso" & "*" & "Giocatore non iscritto;"
										Else
											Semaforo1 = Rec("Semaforo1").Value & "*" & Rec("Titolo1").Value & ";"
										End If
										'If Rec("Semaforo2").Value Is DBNull.Value Or "" & Rec("Semaforo2").Value = "" Then
										'	Semaforo2 = "rosso" & "*" & "Pagamento non completo;"
										'Else
										'Semaforo2 = Rec("Semaforo2").Value & "*" & Rec("Titolo2").Value & ";"
										'End If
										Semaforo2 = "*;"
										If Rec("Semaforo3").Value = "" Then
											Semaforo3 = "rosso" & "*" & "Nessuna firma validata;"
										Else
											Semaforo3 = Rec("Semaforo3").Value & "*" & Rec("Titolo3").Value & ";"
										End If
										If Rec("Semaforo4").Value = "" Then
											Semaforo4 = "rosso" & "*" & "Flag certificato non impostato;"
										Else
											Semaforo4 = Rec("Semaforo4").Value & "*" & Rec("Titolo4").Value & ";"
										End If
										If Rec("Semaforo5").Value = "" Then
											Semaforo5 = "rosso" & "*" & "Nessun elemento kit consegnato;"
										Else
											Semaforo5 = Rec("Semaforo5").Value & "*" & Rec("Titolo5").Value & ";"
										End If

										Dim UtenteGenitore1 As String = ""
										Dim UtenteGenitore2 As String = ""

										If Rec("MailGenitore1").Value <> "" Then
											Sql = "Select * From [Generale].[dbo].[Utenti] Where Utente='" & Rec("MailGenitore1").Value.replace("'", "''") & "'"
											Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
											If TypeOf (Rec2) Is String Then
												Ritorno = Rec2
											Else
												If Not Rec2.Eof() Then
													UtenteGenitore1 = Rec2("Utente").Value
												End If
												'Rec2.Close()
											End If
										End If

										If Rec("MailGenitore2").Value <> "" Then
											Sql = "Select * From [Generale].[dbo].[Utenti] Where Utente='" & Rec("MailGenitore2").Value.replace("'", "''") & "'"
											Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
											If TypeOf (Rec2) Is String Then
												Ritorno = Rec2
											Else
												If Not Rec2.Eof() Then
													UtenteGenitore2 = Rec2("Utente").Value
												End If
												'Rec2.Close()
											End If
										End If

										Ritorno &= Rec("idGiocatore").Value.ToString & ";"
										Ritorno &= Rec("idR").Value.ToString & ";"
										Ritorno &= Rec("Cognome").Value.ToString.Trim & ";"
										Ritorno &= Rec("Nome").Value.ToString.Trim & ";"
										Ritorno &= Rec("Descrizione").Value.ToString.Trim & ";"
										Ritorno &= Rec("EMail").Value.ToString.Trim & ";"
										Ritorno &= Rec("Telefono").Value.ToString.Trim & ";"
										Ritorno &= Rec("Soprannome").Value.ToString.Trim & ";"
										Ritorno &= Rec("DataDiNascita").Value.ToString & ";"
										Ritorno &= Rec("Indirizzo").Value.ToString.Trim & ";"
										Ritorno &= Rec("CodFiscale").Value.ToString.Trim & ";"
										Ritorno &= Rec("Maschio").Value.ToString.Trim & ";"
										Ritorno &= Rec("Citta").Value.ToString.Trim & ";"
										Ritorno &= Rec("Matricola").Value.ToString.Trim & ";"
										Ritorno &= Rec("NumeroMaglia").Value.ToString.Trim & ";"
										Ritorno &= Rec("idCategoria").Value.ToString & ";"
										Ritorno &= Rec("idCategoria2").Value.ToString & ";"
										Ritorno &= Rec("Categoria2").Value.ToString & ";"
										Ritorno &= Rec("idCategoria3").Value.ToString & ";"
										Ritorno &= Rec("Categoria3").Value.ToString & ";"
										Ritorno &= Rec("Categoria1").Value.ToString & ";"
										Ritorno &= Rec("Categorie").Value.ToString & ";"
										Ritorno &= Rec("RapportoCompleto").Value.ToString & ";"
										Ritorno &= Rec("idTaglia").Value.ToString & ";"
										Ritorno &= Semaforo1
										Ritorno &= Semaforo2
										Ritorno &= Semaforo3
										Ritorno &= Semaforo4
										Ritorno &= Semaforo5
										Ritorno &= Rec("idTipologiaKit").Value.ToString & ";"
										Ritorno &= Rec("Cap").Value.ToString & ";"
										Ritorno &= Rec("CittaNascita").Value.ToString & ";"
										Ritorno &= Rec("Maggiorenne").Value.ToString & ";"
										Ritorno &= Rec("CodiceTessera").Value.ToString & ";"
										Ritorno &= UtenteGenitore1 & ";"
										Ritorno &= UtenteGenitore2 & ";"
										Ritorno &= Rec("AmministratiPadre").Value & ";"
										Ritorno &= Rec("AmministratiMadre").Value & ";"

										Dim CeGreenPass As String = "N"
										'Dim PathGP As String = PathBaseMultimedia & NomeSquadra.Replace(" ", "_") & "\GreenPass\" & Rec("idGiocatore").Value & "\GreenPass.kgb"
										Dim path As String = ""
										'If ControllaEsistenzaFile(PathGP) Then
										'	CeGreenPass = "S"
										'	path = pathHTTPMultimedia & NomeSquadra.Replace(" ", "_") & "/GreenPass/" & Rec("idGiocatore").Value & "/GreenPass.kgb"
										'End If

										Ritorno &= CeGreenPass & ";"
										Ritorno &= path & ";"
										Ritorno &= "§"

										Rec.MoveNext()
									Loop
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
						End Try
					End If
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaGreenPass(Squadra As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ": " & Conn
			Else
				Dim Sql As String
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Dim gf As New GestioneFilesDirectory
				Dim Rec As Object
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
					Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
					Dim P() As String = paths.Split(";")
					If Strings.Right(P(0), 1) <> "\" Then
						P(0) &= "\"
					End If
					Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
					If Strings.Right(P(2), 1) <> "/" Then
						P(2) &= "/"
					End If
					Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
					Dim PathBaseMultimedia As String = P(3) ' pathMultimedia.Replace("Allegati", "Multimedia")
					Dim PathGP As String = PathBaseMultimedia & NomeSquadra.Replace(" ", "_") & "\GreenPass\" & idGiocatore & "\GreenPass.kgb"
					Try
						File.Delete(PathGP)
						Ritorno = "*"

						Ritorno = EliminaStringaGreenPass(Squadra, idGiocatore)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If
				gf = Nothing

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornoStringaGreenPass(Squadra As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ": " & Conn
			Else
				Dim Sql As String
				Dim Ok As Boolean = True
				Dim Rec As Object

				Sql = "Select * From GreenPassDati Where idGiocatore=" & idGiocatore
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Return Ritorno
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessuna stringa rilevata"
					Else
						Ritorno = Rec("Stringa").Value
					End If
				End If
				Rec.Close()

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ScriveStringaGreenPass(Squadra As String, idGiocatore As String, StringaGP As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ": " & Conn
			Else
				Dim Sql As String
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Sql = "Delete From GreenPassDati Where idGiocatore=" & idGiocatore
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Ritorno <> "OK" Then
					Ok = False
				End If

				If Ok Then
					Sql = "Insert Into GreenPassDati Values (" &
						" " & idGiocatore & ", " &
						"'" & StringaGP & "'" &
						")"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno <> "OK" Then
						Ok = False
					End If
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
	Public Function EliminaStringaGreenPass(Squadra As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ": " & Conn
			Else
				Dim Sql As String
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Sql = "Delete From GreenPassDati Where idGiocatore=" & idGiocatore
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Ritorno <> "OK" Then
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
	Public Function AggiornaSemafori(Squadra As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ": " & Conn
			Else
				Dim Sql As String
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "---------------------------------")
				Ritorno = CalcolaSemafori(Conn, Connessione, Squadra, idGiocatore)
				If Ritorno <> "OK" Then
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Ritorno: " & Ritorno)
					Ok = False
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Commit")
					Sql = "commit"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Rollback")
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "---------------------------------")

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	Private Function CalcolaSemafori(Conn As Object, Connessione As String, Squadra As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGen As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If ConnessioneGen = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = New clsGestioneDB(Squadra)

			Dim Rec2 As Object
			Dim Sql As String = ""
			Dim Semaforo1 As String = "" : Dim Titolo1 As String = ""
			Dim Semaforo2 As String = "" : Dim Titolo2 As String = ""
			Dim Semaforo3 As String = "" : Dim Titolo3 As String = ""
			Dim Semaforo4 As String = "" : Dim Titolo4 As String = ""
			Dim Semaforo5 As String = "" : Dim Titolo5 As String = ""
			Dim NomeSquadra As String = ""
			Dim iscrFirmaEntrambi As String = ""

			Dim c() As String = Squadra.Split("_")
			Dim Anno As String = Str(Val(c(0))).Trim
			Dim codSquadra As String = c(1)

			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Codice Squadra: " & codSquadra)

			Sql = "Select NomeSquadra, Descrizione, iscrFirmaEntrambi From Anni Where idAnno = " & Anno
			Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
			If TypeOf (Rec2) Is String Then
				Ritorno = Rec2
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
				Return Ritorno
			Else
				If Rec2.Eof() Then
					Ritorno = StringaErrore & " Nessuna squadra rilevata"
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
					Return Ritorno
				Else
					NomeSquadra = "" & Rec2("NomeSquadra").Value
					iscrFirmaEntrambi = "" & Rec2("iscrFirmaEntrambi").Value
				End If
			End If
			Rec2.Close()
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Iscrizione firma entrambi: " & iscrFirmaEntrambi)

			' Semaforo 1: Iscrizione
			Sql = "Select * From Giocatori Where idGiocatore=" & idGiocatore & " "
			Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
			If TypeOf (Rec2) Is String Then
				Ritorno = Rec2
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
				Return Ritorno
			Else
				If Rec2.Eof() Then
					Ritorno = StringaErrore & " Nessun giocatore rilevato"
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
					Return Ritorno
				Else
					Semaforo1 = IIf("" & Rec2("RapportoCompleto").Value = "S", "verde", "rosso")
					Titolo1 = IIf("" & Rec2("RapportoCompleto").Value = "S", "Giocatore iscritto", "Giocatore non iscritto")

					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Rapporto Completo: " & Rec2("RapportoCompleto").Value)
				End If
			End If
			Rec2.Close()

			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Semaforo 1: " & Semaforo1 & " -> " & Titolo1)

			'' Semaforo 2: Pagamenti
			'Sql = "Select Sum(Pagamento) As Pagato, TotalePagamento As Somma " &
			'		"From GiocatoriPagamenti A Left Join GiocatoriDettaglio B On A.idAnno = B.idAnno And A.idGiocatore = B.idGiocatore " &
			'		"Where A.idAnno = " & Anno & " And A.idGiocatore = " & idGiocatore & " " &
			'		"Group By TotalePagamento"
			'Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
			'If TypeOf (Rec2) Is String Then
			'	Ritorno = Rec2
			'	Return Ritorno
			'Else
			'	If Not Rec2.Eof() Then
			'		Semaforo2 = IIf(Rec2("Pagato").Value >= Rec2("Somma").Value, "verde", "giallo")
			'		Titolo2 = IIf(Rec2("Pagato").Value >= Rec2("Somma").Value, "Pagamento completo", "Pagamento parziale")
			'	Else
			'		Semaforo2 = "rosso"
			'		Titolo2 = "Pagamento non completo"
			'	End If
			'	Rec2.Close()
			'End If

			' Semaforo 3: Firme
			Dim GenitoriSeparati As Boolean = False
			Dim AffidamentoCongiunto As Boolean = False
			Dim Maggiorenne As Boolean = False
			Dim idTutore As String = "M"
			Dim AbilitaFirmaGenitore1 As String = ""
			Dim AbilitaFirmaGenitore2 As String = ""
			Dim AbilitaFirmaGenitore3 As String = ""
			Dim FirmaAnalogicaGenitore1 As String = ""
			Dim FirmaAnalogicaGenitore2 As String = ""
			Dim FirmaAnalogicaGenitore3 As String = ""
			Dim quanteFirme As Integer = -1
			If iscrFirmaEntrambi = "S" Then
				quanteFirme = 2
			Else
				quanteFirme = 1
			End If

			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Quante Firme per validazione: " & quanteFirme)

			Sql = "Select * From GiocatoriDettaglio Where idGiocatore=" & idGiocatore & " "
			Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
			If TypeOf (Rec2) Is String Then
				Ritorno = Rec2
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
				Return Ritorno
			Else
				If Not Rec2.Eof() Then
					If "" & Rec2("GenitoriSeparati").Value = "S" Then
						GenitoriSeparati = True
					Else
						GenitoriSeparati = False
					End If

					Maggiorenne = IIf("" & Rec2("Maggiorenne").Value = "S", True, False)
					AffidamentoCongiunto = IIf("" & Rec2("AffidamentoCongiunto").Value = "S", True, False)
					idTutore = "" & Rec2("idTutore").Value
					AbilitaFirmaGenitore1 = "" & Rec2("AbilitaFirmaGenitore1").Value
					AbilitaFirmaGenitore2 = "" & Rec2("AbilitaFirmaGenitore2").Value
					AbilitaFirmaGenitore3 = "" & Rec2("AbilitaFirmaGenitore3").Value
					FirmaAnalogicaGenitore1 = "" & Rec2("FirmaAnalogicaGenitore1").Value
					FirmaAnalogicaGenitore2 = "" & Rec2("FirmaAnalogicaGenitore2").Value
					FirmaAnalogicaGenitore3 = "" & Rec2("FirmaAnalogicaGenitore3").Value

				End If
				Rec2.Close()
			End If

			Dim gf As New GestioneFilesDirectory
			Dim pp As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
			pp = pp.Trim()
			pp = pp.Replace(vbCrLf, "")
			If Strings.Right(pp, 1) <> "\" Then
				pp &= "\"
			End If
			Dim Percorso As String = pp
			Dim q As Integer = 0

			Dim FirmaPresente1 As Boolean = False
			Dim FirmaPresente2 As Boolean = False
			Dim FirmaPresente3 As Boolean = False
			Dim FirmaValidata1 As Boolean = False
			Dim FirmaValidata2 As Boolean = False
			Dim FirmaValidata3 As Boolean = False

			Dim Validate As Integer = 0

			'Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_1.kgb"
			'Dim path2 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_2.kgb"
			'Dim path3 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_3.kgb"

			Dim NumeroFirme As Integer = 2

			Sql = "Select * From NumeroFirme Where idSquadra=" & Val(codSquadra)
			Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGen)
			If TypeOf (Rec2) Is String Then
				Ritorno = Rec2
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
				Return Ritorno
			Else
				If Not Rec2.Eof() Then
					NumeroFirme = Rec2("NumeroFirme").Value
				End If
				Rec2.Close
			End If
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Numero firme: " & NumeroFirme)

			If AbilitaFirmaGenitore1 = "S" Then
				'Firma elettronica attiva genitore 1
				'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, 1, "N") Then
				'	FirmaPresente1 = True
				'	q += 1

				'	Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=1"
				'	Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				'	If TypeOf (Rec2) Is String Then
				'		Ritorno = Rec2
				'		Return Ritorno
				'	Else
				'		If Not Rec2.Eof() Then
				'			If "" & Rec2("Validazione").Value <> "" Then
				'				FirmaValidata1 = True
				'				Validate += 1
				'			End If
				'		End If
				'		Rec2.Close()
				'	End If
				'End If
				Sql = "Select Count(*) As QuanteFirme From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=1 And QualeFirma < 10 And Validazione Is Not Null And Validazione <> ''"
				Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec2) Is String Then
					Ritorno = Rec2
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
					Return Ritorno
				Else
					If Not Rec2.Eof() Then
						If NumeroFirme = Rec2("QuanteFirme").Value Then
							FirmaValidata1 = True
							Validate += 1

							ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Numero firme inserite id 1: " & Rec2("QuanteFirme").Value)
						End If
					End If
					Rec2.Close()
				End If
			Else
				If FirmaAnalogicaGenitore1 = "S" Then
					FirmaPresente1 = True
					FirmaValidata1 = True
					Validate += 1
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma analogica id 1: " & FirmaAnalogicaGenitore1)
				Else
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma analogica id 1: " & FirmaAnalogicaGenitore1 & ". Qui non è gestito")
				End If
			End If

			If AbilitaFirmaGenitore2 = "S" Then
				'Firma elettronica attiva genitore 2
				'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, 2, "N") Then
				'	FirmaPresente2 = True
				'	q += 1

				'	Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=2"
				'	Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				'	If TypeOf (Rec2) Is String Then
				'		Ritorno = Rec2
				'		Return Ritorno
				'	Else
				'		If Not Rec2.Eof() Then
				'			If "" & Rec2("Validazione").Value <> "" Then
				'				FirmaValidata2 = True
				'				Validate += 1
				'			End If
				'		End If
				'		Rec2.Close()
				'	End If
				'End If
				Sql = "Select Count(*) As QuanteFirme From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=2 And QualeFirma < 10 And Validazione Is Not Null And Validazione <> ''"
				Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec2) Is String Then
					Ritorno = Rec2
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
					Return Ritorno
				Else
					If Not Rec2.Eof() Then
						If NumeroFirme = Rec2("QUanteFirme").Value Then
							FirmaValidata2 = True
							Validate += 1

							ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Numero firme inserite id 2: " & Rec2("QuanteFirme").Value)
						End If
					End If
					Rec2.Close()
				End If
			Else
				If FirmaAnalogicaGenitore2 = "S" Then
					FirmaPresente2 = True
					FirmaValidata2 = True
					Validate += 1
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma analogica id 2: " & FirmaAnalogicaGenitore2)
				Else
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma analogica id 2: " & FirmaAnalogicaGenitore2 & ". Qui non è gestito")
				End If
			End If

			If AbilitaFirmaGenitore3 = "S" Then
				'Firma elettronica attiva giocatore
				'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, 3, "N") Then
				'	FirmaPresente3 = True
				'	q += 1

				'	Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=3"
				'	Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				'	If TypeOf (Rec2) Is String Then
				'		Ritorno = Rec2
				'		Return Ritorno
				'	Else
				'		If Not Rec2.Eof() Then
				'			If "" & Rec2("Validazione").Value <> "" Then
				'				FirmaValidata3 = True
				'				Validate += 1
				'			End If
				'		End If
				'		Rec2.Close()
				'	End If
				'End If
				Sql = "Select Count(*) As QuanteFirme From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=3 And QualeFirma < 10 And Validazione Is Not Null And Validazione <> ''"
				Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec2) Is String Then
					Ritorno = Rec2
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
					Return Ritorno
				Else
					If Not Rec2.Eof() Then
						If NumeroFirme = Rec2("QUanteFirme").Value Then
							FirmaValidata3 = True
							Validate += 1

							ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Numero firme inserite id 3: " & Rec2("QuanteFirme").Value)
						End If
					End If
					Rec2.Close()
				End If
			Else
				If FirmaAnalogicaGenitore3 = "S" Then
					FirmaPresente3 = True
					FirmaValidata3 = True
					Validate += 1
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma analogica id 3: " & FirmaAnalogicaGenitore3)
				Else
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma analogica id 3: " & FirmaAnalogicaGenitore3 & ". Qui non è gestito")
				End If
			End If

			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Maggiorenne: " & Maggiorenne)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma Validata 1: " & FirmaValidata1 & " - Firma Presente 1: " & FirmaPresente1)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma Validata 2: " & FirmaValidata2 & " - Firma Presente 2: " & FirmaPresente2)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Firma Validata 3: " & FirmaValidata3 & " - Firma Presente 3: " & FirmaPresente3)

			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Maggiorenne: " & Maggiorenne)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "AffidamentoCongiunto: " & AffidamentoCongiunto)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "idTutore: " & idTutore)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "AbilitaFirmaGenitore1: " & AbilitaFirmaGenitore1)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "AbilitaFirmaGenitore2: " & AbilitaFirmaGenitore2)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "AbilitaFirmaGenitore3: " & AbilitaFirmaGenitore3)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "FirmaAnalogicaGenitore1: " & FirmaAnalogicaGenitore1)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "FirmaAnalogicaGenitore2: " & FirmaAnalogicaGenitore2)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "FirmaAnalogicaGenitore3: " & FirmaAnalogicaGenitore3)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "GenitoriSeparati: " & GenitoriSeparati)

			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Validate: " & Validate)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Quante Firme: " & quanteFirme)

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
								If Validate < quanteFirme Then
									Semaforo3 = "giallo"
									Titolo3 = "Firme non validate (" & Validate & "/" & quanteFirme & ")"
								Else
									Semaforo3 = "verde"
									Titolo3 = "Tutte le firme validate (" & quanteFirme & ")"
								End If
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
									If Validate < quanteFirme Then
										Semaforo3 = "giallo"
										Titolo3 = "Firme non validate (" & Validate & "/" & quanteFirme & ")"
									Else
										Semaforo3 = "verde"
										Titolo3 = "Tutte le firme validate (" & quanteFirme & ")"
									End If
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
									If Validate < quanteFirme Then
										Semaforo3 = "giallo"
										Titolo3 = "Firme non validate (" & Validate & "/" & quanteFirme & ")"
									Else
										Semaforo3 = "verde"
										Titolo3 = "Tutte le firme validate (" & quanteFirme & ")"
									End If
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
							If Validate < quanteFirme Then
								Semaforo3 = "giallo"
								Titolo3 = "Firme non validate (" & Validate & "/" & quanteFirme & ")"
							Else
								Semaforo3 = "verde"
								Titolo3 = "Tutte le firme validate (" & quanteFirme & ")"
							End If
						Else
							Semaforo3 = "rosso"
							Titolo3 = "Nessuna firma validata"
						End If
					End If
				End If
			End If

			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Semaforo3: " & Semaforo3)
			ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Titolo Semaforo3: " & Titolo3)

			'Semaforo 4: Certificato
			Try
				If TipoDB = "SQLSERVER" Then
					Sql = "Select CertificatoMedico, IsNull(ScadenzaCertificatoMedico,'') As ScadenzaCertificatoMedico From GiocatoriDettaglio " &
					"Where idGiocatore = " & idGiocatore
				Else
					Sql = "Select CertificatoMedico, Coalesce(ScadenzaCertificatoMedico, '') As ScadenzaCertificatoMedico From GiocatoriDettaglio " &
					"Where idGiocatore = " & idGiocatore
				End If
				Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec2) Is String Then
					Ritorno = Rec2
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
					Return Ritorno
				Else
					If Not Rec2.Eof() Then
						ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "" & Rec2("CertificatoMedico").Value)
						If ("" & Rec2("CertificatoMedico").Value) = "" Or ("" & Rec2("CertificatoMedico").Value) = "N" Then
							Semaforo4 = "rosso"
							Titolo4 = "Flag certificato non impostato"
						Else
							If Rec2("ScadenzaCertificatoMedico").Value = "" Then
								If "" & Rec2("CertificatoMedico").Value = "S" Then
									Semaforo4 = "giallo"
									Titolo4 = "Certificato presente, Scadenza no"
								Else
									Semaforo4 = "rosso"
									Titolo4 = "Nessun certificato e data presenti"
								End If
							Else
								If "" & Rec2("ScadenzaCertificatoMedico").Value = "" Then
									If "" & Rec2("CertificatoMedico").Value = "S" Then
										Semaforo4 = "giallo"
										Titolo4 = "Certificato presente, Scadenza no"
									Else
										Semaforo4 = "rosso"
										Titolo4 = "Nessun certificato e data presenti"
									End If
								Else
									Dim dd As String = "" & Rec2("ScadenzaCertificatoMedico").Value
									If dd <> "" Then
										Try
											If dd.Contains("-") Then
												Dim D() As String = dd.Split("-")
												Dim dat As Date = New Date(Val(D(0)), Val(D(1)), Val(D(2))) ' Convert.ToDateTime(D(2) & "/" & D(1) & "/" & D(0))

												ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Scadenza: " & dd & " -> " & Val(D(0)) & "/" & Val(D(1)) & "/" & Val(D(2)))

												'Dim Scadenza As DateTime = Convert.ToDateTime("" & Rec2("ScadenzaCertificatoMedico").Value)
												Dim GiorniAllaScadenza As Integer = DateAndTime.DateDiff(DateInterval.Day, Now, dat, ) + 1

												ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Giorni alla scadenza: " & GiorniAllaScadenza)

												If "" & Rec2("CertificatoMedico").Value = "S" And dat > Now.AddDays(-1) Then
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
											Else
												Semaforo4 = "rosso"
												Titolo4 = "Certificato presente ma con data non valida" ' (" & dd & ")"
											End If
										Catch ex As Exception
											Semaforo4 = "rosso"
											Titolo4 = "Certificato presente ma conversione data non valida" ' (" & dd & ")"
										End Try
									Else
										Semaforo4 = "rosso"
										Titolo4 = "Data Certificato non presente"
									End If
								End If
							End If
						End If
					Else
						Semaforo4 = "rosso"
						Titolo4 = "Nessun certificato e data presenti"
					End If
					Rec2.Close()
				End If
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Semaforo 4: " & Semaforo4)
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Titolo Semaforo 4: " & Titolo4)
			Catch ex As Exception
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Errore controllo certificato: " & ex.Message)
			End Try

			' Semaforo 5: KIT
			Try
				Sql = "Select C.Descrizione, QuantitaConsegnata, Quantita From KitGiocatori A " &
					"Left Join KitTipologie B On A.idTipoKit = B.idTipoKit " &
					"Left Join KitElementi C On A.idElemento = C.idElemento " &
					"Left Join KitComposizione D On D.idAnno = " & Anno & " And A.idTipoKit = B.idTipoKit And A.idElemento = C.idElemento And A.idTipoKit = D.idTipoKit  And A.idElemento = D.idElemento " &
					"Where idGiocatore = " & idGiocatore & " And B.Eliminato = 'N' And C.Eliminato = 'N' And D.Eliminato = 'N'"
				Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec2) Is String Then
					Ritorno = Rec2
					ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
					Return Ritorno
				Else
					If Rec2.Eof() Then
						Semaforo5 = "rosso"
						Titolo5 = "Nessun elemento kit consegnato"
					Else
						Dim Tutto As Boolean = True
						Dim Qualcosa As Boolean = False

						Do Until Rec2.Eof()
							If Val(Rec2("QuantitaConsegnata").Value) > 0 Then
								Qualcosa = True
								If Val(Rec2("QuantitaConsegnata").Value) < Val(Rec2("Quantita").Value) Then
									Tutto = False
									Exit Do
								End If
							Else
								If Val(Rec2("Quantita").Value) > 0 Then
									Tutto = False
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
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Semaforo 5" & Semaforo5)
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Titolo Semaforo 5" & Titolo5)
			Catch ex As Exception
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", "Errore controllo kit: " & ex.Message)
			End Try

			Sql = "Delete From GiocatoriSemafori Where idGiocatore=" & idGiocatore
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			If Ritorno <> "OK" Then
				ScriveLog(Server.MapPath("."), Squadra, "CalcolaSemafori", Ritorno)
				Return Ritorno
			End If

			Sql = "Insert Into GiocatoriSemafori Values (" &
				" " & idGiocatore & ", " &
				"'" & Semaforo1.Replace("'", "''") & "', " &
				"'" & Titolo1.Replace("'", "''") & "', " &
				"'" & Semaforo2.Replace("'", "''") & "', " &
				"'" & Titolo2.Replace("'", "''") & "', " &
				"'" & Semaforo3.Replace("'", "''") & "', " &
				"'" & Titolo3.Replace("'", "''") & "', " &
				"'" & Semaforo4.Replace("'", "''") & "', " &
				"'" & Titolo4.Replace("'", "''") & "', " &
				"'" & Semaforo5.Replace("'", "''") & "', " &
				"'" & Titolo5.Replace("'", "''") & "' " &
				")"
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
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
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
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
	Public Function InviaIscrizione(Squadra As String, idAnno As String, idGiocatore As String, Mittente As String, Destinatario As String, NomeSquadra As String, Giocatore As String) As String
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
				Dim fileDaInviarePDF As String = ""

				ScriveLog(Server.MapPath("."), Squadra, "InvioIscrizione", "-------------------------------")
				Dim imm As New wsImmagini
				fileDaInviarePDF = imm.RitornaAllegatoDB2(Squadra, "iscrizioni", idGiocatore, 1)
				ScriveLog(Server.MapPath("."), Squadra, "InvioIscrizione", "File da inviare: " & fileDaInviarePDF)
				If Ritorno.Contains("ERROR:") Then
					If Ritorno.Contains("Nessun dato ricevuto") Then
					Else
						Ok = False
					End If
				End If

				If Ok Then
					Dim filesDaAllegare() As String = {Server.MapPath(".") & "\" & fileDaInviarePDF}
					Dim Corpo As String = Giocatore & " e' stato iscritto. In allegato si inoltra copia del modulo d'iscrizione.<br /><br />La segreteria " & NomeSquadra
					Dim Oggetto As String = NomeSquadra & ": Iscrizione di " & Giocatore
					Dim m As New mail

					ScriveLog(Server.MapPath("."), Squadra, "InvioIscrizione", "Invio mail")
					Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, Corpo, Destinatario, filesDaAllegare)
					If Ritorno = "*" Then
						Dim gf As New GestioneFilesDirectory
						gf.EliminaFileFisico(fileDaInviarePDF)
						gf = Nothing
					End If
					ScriveLog(Server.MapPath("."), Squadra, "InvioIscrizione", "Ritorno mail: " & Ritorno)
				End If
				ScriveLog(Server.MapPath("."), Squadra, "InvioIscrizione", "-------------------------------")
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Update Giocatori Set " &
							"RapportoCompleto='" & Replace(RapportoCompleto, "'", "''") & "' " &
							"Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
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
	Public Function SalvaDettaglioGiocatore(Squadra As String, idAnno As String, idGiocatore As String, Genitore1 As String,
											Genitore2 As String, FirmaGenitore1 As String, FirmaGenitore2 As String, CertificatoMedico As String,
											TotalePagamento As String, TelefonoGenitore1 As String, TelefonoGenitore2 As String,
											ScadenzaCertificatoMedico As String, MailGenitore1 As String, MailGenitore2 As String, FirmaGenitore3 As String, MailGenitore3 As String,
											DataDiNascita1 As String, CittaNascita1 As String, CodFiscale1 As String, Citta1 As String, Cap1 As String, Indirizzo1 As String,
											DataDiNascita2 As String, CittaNascita2 As String, CodFiscale2 As String, Citta2 As String, Cap2 As String, Indirizzo2 As String,
											GenitoriSeparati As String, AffidamentoCongiunto As String, AbilitaFirmaGenitore1 As String, AbilitaFirmaGenitore2 As String,
											AbilitaFirmaGenitore3 As String, FirmaAnalogicaGenitore1 As String, FirmaAnalogicaGenitore2 As String, FirmaAnalogicaGenitore3 As String,
											idTutore As String, idQuota As String, FirmaGenitore4 As String, AbilitaFirmaGenitore4 As String, FirmaAnalogicaGenitore4 As String,
											NoteKit As String, Sconto As String) As String
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

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

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
							"MailGenitore1='" & MailGenitore1.Replace("'", "''") & "', " &
							"MailGenitore2='" & MailGenitore2.Replace("'", "''") & "', " &
							"FirmaGenitore3='" & Replace(FirmaGenitore3, "'", "''") & "', " &
							"FirmaGenitore4='" & Replace(FirmaGenitore4, "'", "''") & "', " &
							"MailGenitore3='" & MailGenitore3.Replace("'", "''") & "', " &
							"DataDiNascita1='" & DataDiNascita1 & "', " &
							"CittaNascita1='" & CittaNascita1.Replace("'", "''") & "', " &
							"CodFiscale1='" & CodFiscale1 & "', " &
							"Citta1='" & Citta1.Replace("'", "''") & "', " &
							"Cap1='" & Cap1 & "', " &
							"Indirizzo1='" & Indirizzo1.Replace("'", "''") & "', " &
							"DataDiNascita2='" & DataDiNascita2 & "', " &
							"CittaNascita2='" & CittaNascita2.Replace("'", "''") & "', " &
							"CodFiscale2='" & CodFiscale2 & "', " &
							"Citta2='" & Citta2.Replace("'", "''") & "', " &
							"Cap2='" & Cap2 & "', " &
							"Indirizzo2='" & Indirizzo2.Replace("'", "''") & "', " &
							"GenitoriSeparati='" & GenitoriSeparati & "', " &
							"AffidamentoCongiunto='" & AffidamentoCongiunto & "', " &
							"AbilitaFirmaGenitore1='" & AbilitaFirmaGenitore1 & "', " &
							"AbilitaFirmaGenitore2='" & AbilitaFirmaGenitore2 & "', " &
							"AbilitaFirmaGenitore3='" & AbilitaFirmaGenitore3 & "', " &
							"AbilitaFirmaGenitore4='" & AbilitaFirmaGenitore4 & "', " &
							"FirmaAnalogicaGenitore1='" & FirmaAnalogicaGenitore1 & "', " &
							"FirmaAnalogicaGenitore2='" & FirmaAnalogicaGenitore2 & "', " &
							"FirmaAnalogicaGenitore3='" & FirmaAnalogicaGenitore3 & "', " &
							"FirmaAnalogicaGenitore4='" & FirmaAnalogicaGenitore4 & "', " &
							"idTutore='" & idTutore & "', " &
							"Sconto=" & Sconto & ", " &
							"idQuota='" & idQuota & "' " &
							"Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						Sql = "Delete From KitNote Where idGiocatore=" & idGiocatore
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If
					If Ok Then
						Sql = "Insert Into KitNote Values(" & idGiocatore & ", '" & NoteKit.Replace("'", "''") & "')"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						Sql = "Delete From GiocatoriMails Where idGiocatore=" & idGiocatore
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						'Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
						'If TypeOf (Rec) Is String Then
						'	Ritorno = Rec
						'Else
						'	If Rec.Eof() Then
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
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							Else
							End If
							'End If
						End If
						'End If
						'Rec.Close()
					End If
					'End If
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
	Public Function TornaDatiGiocatore(NumeroTessera As String, Squadra As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim c2() As String = Squadra.Split("_")
		Dim Anno As String = Str(Val(c2(0))).Trim
		Dim codSquadra As String = c2(1)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From Giocatori Where idGiocatore=" & idGiocatore

				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = "ERROR: Nessun giocatore rilevato"
					Else
						Ritorno = Rec("Cognome").Value & ";"
						Ritorno &= Rec("Nome").Value & ";"
						Ritorno &= Rec("CodFiscale").Value & ";"
						Dim Campi() As String = Rec("Categorie").Value.split("-")
						Rec.Close()

						Sql = "Select * From Anni Where idAnno=" & Anno
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ritorno = "ERROR: Nessun giocatore rilevato"
							Else
								Dim Percentuale As String = Rec("PercCashBack").Value
								Rec.Close()

								Dim Categorie As String = ""

								For Each c As String In Campi
									If c <> "" Then
										Sql = "Select * From Categorie Where idCategoria=" & c
										Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
										If TypeOf (Rec) Is String Then
										Else
											If Not Rec.Eof() Then
												If Not Categorie.Contains(Rec("Descrizione").Value) Then
													Categorie &= Rec("Descrizione").Value & "*"
												End If
											End If
										End If
									End If
								Next

								Ritorno &= Categorie & ";"

								If TipoDB = "SQLSERVER" Then
									Sql = "Select IsNull(Sum(Importo),0) From [Generale].[dbo].[TessereNFC] Where NumeroTessera='" & NumeroTessera & "'" ' CodSquadra='" & Squadra & "' And idGiocatore=" & idGiocatore
								Else
									Sql = "Select Coalesce(Sum(Importo),0) From [Generale].[dbo].[TessereNFC] Where NumeroTessera='" & NumeroTessera & "'" ' CodSquadra='" & Squadra & "' And idGiocatore=" & idGiocatore
								End If
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									Dim Saldo As String = ""

									' Return Rec(0).Value & ";" & Percentuale & ""

									'If Rec(0).Value Is DBNull.Value Then
									'	Saldo = "€ 0"
									'Else
									Saldo = "€ " & Rec(0).Value
									'End If

									Ritorno &= Saldo & ";"

									Ritorno &= Percentuale & ";"
									Dim ss As Single = Rec(0).Value
									Dim pp As Integer = Val(Percentuale.Replace(",", "."))
									Dim cb As Single = CInt((ss * pp / 100) * 100) / 100

									'Return ss & ";" & pp & ";" & cb

									Ritorno &= cb

									Rec.Close()
								End If
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TornaDettaglioGiocatore(Squadra As String, idAnno As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim ConnessioneGen As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim ConnGen As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Rec2 As Object
				Dim Sql As String = ""
				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)

				Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Dim NomeSquadra As String = ""
					Dim Descrizione As String = ""

					If Rec.Eof() Or "" & Rec("NomeSquadra").Value = "" Then
						Descrizione = "" & Rec("Descrizione").Value

						Sql = "Select * From Squadre Where idSquadra=" & Val(codSquadra)
						Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGen)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Not Rec.Eof() Then
								NomeSquadra = Rec("Descrizione").Value

								Sql = "Update Anni Set NomeSquadra='" & NomeSquadra & "'"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							End If
						End If
					Else
						NomeSquadra = "" & Rec("NomeSquadra").Value
						Descrizione = "" & Rec("Descrizione").Value
					End If
					Rec.Close()

					If NomeSquadra = "" Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						Dim ratePagate As String = ":"

						Sql = "Select Distinct B.idRata From GiocatoriDettaglio A " &
							"Left Join GiocatoriPagamenti B On A.idGiocatore = B.idGiocatore " &
							"Where A.idGiocatore = " & idGiocatore & " " ' And Validato = 'S'"
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof()
								If (("" & Rec("idRata").Value).contains(";")) Then
									Dim rr() As String = Rec("idRata").Value.split(";")

									For Each r As String In rr
										If r <> "" Then
											ratePagate &= r & ":"
										End If
									Next
								Else
									ratePagate &= Rec("idRata").Value & ":"
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()
						End If

						Dim importiManuali As String = ":"

						Sql = "Select B.Progressivo, B.ImportoManuale, B.DescrizioneManuale, B.DataManuale From GiocatoriDettaglio A " &
							"Left Join GiocatoriPagamenti B On A.idGiocatore = B.idGiocatore " &
							"Where A.idGiocatore = " & idGiocatore & " And " & IIf(TipoDB = "SQLSERVER", "CHARINDEX('9999;', B.idRata) > 0", "Instr(B.idRata, '9999;') > 0") & " "
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof()
								importiManuali &= Rec("Progressivo").Value & "%" & Rec("ImportoManuale").Value & "%" & ("" & Rec("DescrizioneManuale").Value).replace(";", "***PV***").replace(":", "***2P***").replace("%", "***PE***") & "%" & Rec("DataManuale").Value & ":"

								Rec.MoveNext()
							Loop
							Rec.Close()
						End If

						Sql = "Select * From GiocatoriDettaglio A " &
							"Left Join KitNote B On A.idGiocatore = B.idGiocatore " &
							"Where A.idAnno=" & Anno & " And A.idGiocatore=" & idGiocatore
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Rec.Close()

								Dim totPagamento As String = "0"

								If TipoDB = "SQLSERVER" Then
									Sql = "Select IsNull(CostoScuolaCalcio,0) As CostoScuolaCalcio From Anni"
								Else
									Sql = "Select Coalesce(CostoScuolaCalcio,0) As CostoScuolaCalcio From Anni"
								End If
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Not Rec.Eof() Then
										'If Not Rec("CostoScuolaCalcio").Value Is DBNull.Value Then
										totPagamento = Rec("CostoScuolaCalcio").Value '.replace(",", ".")
										'Else
										'	totPagamento = 0
										'End If
									End If
									Rec.Close()
								End If

								Sql = "Insert Into GiocatoriDettaglio Values (" &
									" " & idAnno & ", " &
									" " & idGiocatore & ", " &
									"'', " &
									"'', " &
									"'N', " &
									"'N', " &
									"'N', " &
									" " & totPagamento.Replace(",", ".") & ", " &
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
									"'', " &
									"'N', " &
									"'N', " &
									"'S', " &
									"'S', " &
									"'S', " &
									"'N', " &
									"'N', " &
									"'N', " &
									"'M', " &
									"-1, " &
									"'', " &
									"'S', " &
									"'N', " &
									"0" &
									")"
								'idAnno  idGiocatore	Genitore1	Genitore2	FirmaGenitore1	FirmaGenitore2	CertificatoMedico	TotalePagamento	TelefonoGenitore1	TelefonoGenitore2	ScadenzaCertificatoMedico	MailGenitore1	
								'MailGenitore2   FirmaGenitore3	MailGenitore3	DataDiNascita1	CittaNascita1	CodFiscale1	Citta1	Cap1	Indirizzo1	DataDiNascita2	CittaNascita2	CodFiscale2	Citta2	Cap2	Indirizzo2	
								'Maggiorenne GenitoriSeparati	AffidamentoCongiunto	AbilitaFirmaGenitore1	AbilitaFirmaGenitore2	AbilitaFirmaGenitore3	FirmaAnalogicaGenitore1	FirmaAnalogicaGenitore2	FirmaAnalogicaGenitore3	
								'idTutore    idQuota	FirmaGenitore4	AbilitaFirmaGenitore4	FirmaAnalogicaGenitore4 Sconto

								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

								If Not Ritorno.Contains(StringaErrore) Then
									Sql = "Select * From GiocatoriDettaglio A " &
										"Left Join KitNote B On A.idGiocatore = B.idGiocatore " &
										"Where A.idAnno=" & Anno & " And A.idGiocatore=" & idGiocatore
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										If Not Rec.Eof() Then
											Ritorno = Rec("idAnno").Value & ";"
											Ritorno &= Rec("idGiocatore").Value & ";"
											Ritorno &= Rec("Genitore1").Value & ";"
											Ritorno &= Rec("Genitore2").Value & ";"
											Ritorno &= Rec("FirmaGenitore1").Value & ";"
											Ritorno &= Rec("FirmaGenitore2").Value & ";"
											Ritorno &= Rec("CertificatoMedico").Value & ";"
											Ritorno &= Rec("TotalePagamento").Value & ";"
											Ritorno &= Rec("TelefonoGenitore1").Value & ";"
											Ritorno &= Rec("TelefonoGenitore2").Value & ";"
											Ritorno &= Rec("ScadenzaCertificatoMedico").Value & ";"
											Ritorno &= Rec("MailGenitore1").Value & ";"
											Ritorno &= Rec("MailGenitore2").Value & ";"
											Ritorno &= Rec("FirmaGenitore3").Value & ";"
											Ritorno &= Rec("MailGenitore3").Value & ";"
											Ritorno &= Rec("DataDiNascita1").Value & ";"
											Ritorno &= Rec("CittaNasciat1").Value & ";"
											Ritorno &= Rec("CodFiscale1").Value & ";"
											Ritorno &= Rec("Citta1").Value & ";"
											Ritorno &= Rec("Indirizzo1").Value & ";"
											Ritorno &= Rec("DataDiNascita1").Value & ";"
											Ritorno &= Rec("CittaNasciat2").Value & ";"
											Ritorno &= Rec("CodFiscale2").Value & ";"
											Ritorno &= Rec("Citta2").Value & ";"
											Ritorno &= Rec("Cap2").Value & ";"
											Ritorno &= Rec("Indirizzo2").Value & ";"
											Ritorno &= Rec("Maggiorenne").Value & ";"
											Ritorno &= Rec("GenitoriSeparati").Value & ";"
											Ritorno &= Rec("AffidamentoCongiunto").Value & ";"
											Ritorno &= Rec("AbilitaFirmaGenitore1").Value & ";"
											Ritorno &= Rec("AbilitaFirmaGenitore2").Value & ";"
											Ritorno &= Rec("AbilitaFirmaGenitore3").Value & ";"
											Ritorno &= Rec("FirmaAnalogicaGenitore1").Value & ";"
											Ritorno &= Rec("FirmaAnalogicaGenitore2").Value & ";"
											Ritorno &= Rec("FirmaAnalogicaGenitore3").Value & ";"
											Ritorno &= Rec("idTutore").Value & ";"
											Ritorno &= Rec("idQuota").Value & ";"
											Ritorno &= ratePagate & ";"
											Dim n As String = "" & Rec("Note").Value
											Ritorno &= n.Replace(";", "***PV***") & ";"
											Ritorno &= Rec("Sconto").Value & ";"

											Sql = "Select * From Quote Where idQuota=" & Rec("idQuota").Value
											Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
											If Rec2.Eof() Then
												Ritorno &= "Quota non impostata;"
												Ritorno &= "0;"
											Else
												Ritorno &= Rec2("Descrizione").Value.replace(";", "***PV***").replace(":", "***2P***").replace("%", "***PE***") & ";"
												Ritorno &= Rec2("Importo").Value & ";"
											End If
											Rec2.Close()

											Ritorno &= importiManuali & ";"

											If TipoDB = "SQLSERVER" Then
												Sql = "Select IsNull(Max(Progressivo),-1) From QuoteRate Where Attiva='S' And Importo > 0 And idQuota = " & Rec("idQuota").Value
											Else
												Sql = "Select Coalesce(Max(Progressivo),-1) From QuoteRate Where Attiva='S' And Importo > 0 And idQuota = " & Rec("idQuota").Value
											End If
											Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
											'If Rec2(0).Value Is DBNull.Value Then
											'	Ritorno &= "-1;"
											'Else
											Ritorno &= Rec2(0).Value & ";"
											'End If
											Rec2.Close()

											Sql = "Select " & IIf(TipoDB = "SQLSERVER", "ISNULL(Sum(Pagamento),0)", "COALESCE(Sum(Pagamento),0)") & " From GiocatoriPagamenti " &
												"Where idGiocatore = " & Rec("idGiocatore").Value & " And Eliminato = 'N' And Validato = 'S' And idTipoPagamento = 1"
											Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
											'If Rec2(0).Value Is DBNull.Value Then
											'	Ritorno &= "0;"
											'Else
											Ritorno &= Rec2(0).Value & ";"
											'End If
											Rec2.Close()

										End If
										Rec.Close()
									End If
								End If
							Else
								Dim gf As New GestioneFilesDirectory
								Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
								'gf = Nothing
								Percorso = Percorso.Trim()
								If Strings.Right(Percorso, 1) <> "\" Then
									Percorso &= "\"
								End If
								'Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_1.kgb"
								'Dim path2 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_2.kgb"
								'Dim path3 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_3.kgb"
								'Dim path4 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_4.kgb"

								Dim dataFirma1 As String = ""
								Dim dataFirma2 As String = ""
								Dim dataFirma3 As String = ""
								Dim dataFirma4 As String = ""

								Dim NumeroFirme As Integer = 2

								Sql = "Select * From NumeroFirme Where idSquadra=" & Val(codSquadra)
								Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGen)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2
								Else
									If Rec2.Eof() = False Then
										NumeroFirme = "" & Rec2("NumeroFirme").Value
									End If
									Rec2.Close()
								End If

								'gf.CreaDirectoryDaPercorso(Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\")

								Dim firma1(NumeroFirme - 1) As String
								Dim QuanteFirmeValidate1 As Integer = 0
								Dim QuanteFirmeValidate2 As Integer = 0
								Dim QuanteFirmeValidate3 As Integer = 0
								Dim QuanteFirmeValidate4 As Integer = 0

								For i As Integer = 0 To NumeroFirme - 1
									firma1(i) = "N"
								Next
								Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=1 And Validazione Is Not Null And Validazione <> '' Order By QualeFirma"
								Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2
								Else
									Do Until Rec2.Eof
										If Rec2("QualeFirma").Value <= firma1.Length Then
											firma1(Rec2("QualeFirma").Value - 1) = "S"
											dataFirma1 &= "" & Rec2("DataFirma").Value & "%"
											QuanteFirmeValidate1 += 1
										End If

										Rec2.MoveNext
									Loop
									Rec2.Close()
								End If

								Dim firma2(NumeroFirme - 1) As String

								For i As Integer = 0 To NumeroFirme - 1
									firma2(i) = "N"
								Next
								Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=2 And Validazione Is Not Null And Validazione <> '' Order By QualeFirma"
								Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2
								Else
									Do Until Rec2.Eof
										firma2(Rec2("QualeFirma").Value - 1) = "S"
										dataFirma2 &= "" & Rec2("DataFirma").Value & "%"
										QuanteFirmeValidate2 += 1

										Rec2.MoveNext
									Loop
									Rec2.Close()
								End If

								Dim firma3(NumeroFirme - 1) As String

								For i As Integer = 0 To NumeroFirme - 1
									firma3(i) = "N"
								Next
								Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=3 And Validazione Is Not Null And Validazione <> '' Order By QualeFirma"
								Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2
								Else
									Do Until Rec2.Eof
										firma3(Rec2("QualeFirma").Value - 1) = "S"
										dataFirma3 &= "" & Rec2("DataFirma").Value & "%"
										QuanteFirmeValidate3 += 1

										Rec2.MoveNext
									Loop
									Rec2.Close()
								End If

								Dim firma4(NumeroFirme - 1) As String

								For i As Integer = 0 To NumeroFirme - 1
									firma4(i) = "N"
								Next
								Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=4 And Validazione Is Not Null And Validazione <> '' Order By QualeFirma"
								Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2
								Else
									Do Until Rec2.Eof
										firma4(Rec2("QualeFirma").Value - 1) = "S"
										dataFirma4 &= "" & Rec2("DataFirma").Value & "%"
										QuanteFirmeValidate4 += 1

										Rec2.MoveNext
									Loop
									Rec2.Close()
								End If

								Dim sFirma1 As String = ""
								Dim sFirma2 As String = ""
								Dim sFirma3 As String = ""
								Dim sFirma4 As String = ""

								For i As Integer = 0 To firma1.Length - 1
									sFirma1 &= firma1(i) & "%"
								Next
								For i As Integer = 0 To firma2.Length - 1
									sFirma2 &= firma2(i) & "%"
								Next
								For i As Integer = 0 To firma3.Length - 1
									sFirma3 &= firma3(i) & "%"
								Next
								For i As Integer = 0 To firma4.Length - 1
									sFirma4 &= firma4(i) & "%"
								Next

								'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, 1, "N") Then
								'	firma1 = "S"
								'	Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=1"
								'	Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								'	If TypeOf (Rec2) Is String Then
								'		Ritorno = Rec2
								'	Else
								'		If Rec2.Eof() = False Then
								'			dataFirma1 = "" & Rec2("DataFirma").Value
								'		End If
								'		Rec2.Close()
								'	End If
								'End If

								'Dim firma2 As String = "N"
								'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, 2, "N") Then
								'	firma2 = "S"
								'	Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=2"
								'	Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								'	If TypeOf (Rec2) Is String Then
								'		Ritorno = Rec2
								'	Else
								'		If Rec2.Eof() = False Then
								'			dataFirma2 = "" & Rec2("DataFirma").Value
								'		End If
								'		Rec2.Close()
								'	End If
								'End If

								' Return firma2 & "-" & dataFirma2 & " -> " & Sql

								'Dim firma3 As String = "N"
								'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, 3, "N") Then
								'	firma3 = "S"
								'	Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=3"
								'	Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								'	If TypeOf (Rec2) Is String Then
								'		Ritorno = Rec2
								'	Else
								'		If Rec2.Eof() = False Then
								'			dataFirma3 = "" & Rec2("DataFirma").Value
								'		End If
								'		Rec2.Close()
								'	End If
								'End If

								'Dim firma4 As String = "N"
								'If ControllaEsistenzaFile(Server.MapPath("."), Conn, Connessione, idGiocatore, 4, "N") Then
								'	firma4 = "S"
								'	Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=4"
								'	Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								'	If TypeOf (Rec2) Is String Then
								'		Ritorno = Rec2
								'	Else
								'		If Rec2.Eof() = False Then
								'			dataFirma4 = "" & Rec2("DataFirma").Value
								'		End If
								'		Rec2.Close()
								'	End If
								'End If

								Ritorno = Rec("Genitore1").Value & ";"
								Ritorno &= Rec("Genitore2").Value & ";"
								Ritorno &= sfirma1 & ";"
								Ritorno &= sfirma2 & ";"
								Ritorno &= Rec("CertificatoMedico").Value & ";"
								Ritorno &= Rec("TotalePagamento").Value & ";"
								Ritorno &= Rec("TelefonoGenitore1").Value & ";"
								Ritorno &= Rec("TelefonoGenitore2").Value & ";"
								Ritorno &= Rec("ScadenzaCertificatoMedico").Value & ";"
								Ritorno &= Rec("MailGenitore1").Value & ";"
								Ritorno &= Rec("MailGenitore2").Value & ";"
								Ritorno &= sfirma3 & ";"
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
								Ritorno &= Rec("idQuota").Value & ";"
								Ritorno &= ratePagate & ";"

								Ritorno &= dataFirma4 & ";"
								Ritorno &= Rec("AbilitaFirmaGenitore4").Value & ";"
								Ritorno &= Rec("FirmaAnalogicaGenitore4").Value & ";"
								Ritorno &= sfirma4 & ";"
								Dim n As String = "" & Rec("Note").Value
								Ritorno &= n.Replace(";", "***PV***") & ";"
								Ritorno &= Rec("Sconto").Value & ";"

								' If Rec("idQuota").Value Is DBNull.Value Then ***NULL DA SISTEMARE***
								If Rec("idQuota").Value = "<NULL>" Then
									Ritorno &= "Quota non impostata;"
									Ritorno &= "0;"
								Else
									If "" & Rec("idQuota").Value = "" Then
										Ritorno &= "Quota non impostata;"
										Ritorno &= "0;"
									Else
										Sql = "Select * From Quote Where idQuota=" & Rec("idQuota").Value
										Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Rec2.Eof() Then
												Ritorno &= "Quota non impostata;"
												Ritorno &= "0;"
											Else
												Ritorno &= Rec2("Descrizione").Value.replace(";", "***PV***").replace(":", "***2P***").replace("%", "***PE***") & ";"
												Ritorno &= Rec2("Importo").Value & ";"
											End If
										End If
										Rec2.Close()
									End If
								End If

								Ritorno &= importiManuali & ";"

								'If Rec("idQuota").Value Is DBNull.Value Then ***NULL DA SISTEMARE***
								If Rec("idQuota").Value = "<NULL>" Then
									Ritorno &= "-1;"
								Else
									If "" & Rec("idQuota").Value = "" Then
										Ritorno &= "-1;"
									Else
										If TipoDB = "SQLSERVER" Then
											Sql = "Select Isnull(Max(Progressivo),-1) From QuoteRate Where Attiva='S' And Importo > 0 And idQuota = " & Rec("idQuota").Value
										Else
											Sql = "Select Coalesce(Max(Progressivo),-1) From QuoteRate Where Attiva='S' And Importo > 0 And idQuota = " & Rec("idQuota").Value
										End If
										Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
										'If Rec2(0).Value Is DBNull.Value Or Rec2(0).Value = "" Then
										'	Ritorno &= "-1;"
										'Else
										Ritorno &= Rec2(0).Value & ";"
										'End If
										Rec2.Close()
									End If
								End If

								If TipoDB = "SQLSERVER" Then
									Sql = "Select isnull(Sum(Pagamento),0) From GiocatoriPagamenti " &
												"Where idGiocatore = " & idGiocatore & " And Eliminato = 'N' And Validato = 'S' And idTipoPagamento = 1"
								Else
									Sql = "Select Coalesce(Sum(Pagamento),0) From GiocatoriPagamenti " &
												"Where idGiocatore = " & idGiocatore & " And Eliminato = 'N' And Validato = 'S' And idTipoPagamento = 1"
								End If
								Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								'If Rec2(0).Value Is DBNull.Value Then
								'	Ritorno &= "0;"
								'Else
								Ritorno &= Rec2(0).Value & ";"
								'End If
								Rec2.Close()

								Dim wI As New wsImmagini()
								Dim pathPdf As String = wI.RitornaAllegatoDB2(Squadra, "iscrizioni", idGiocatore, 1)
								If pathPdf.Contains("ERROR:") Then
									pathPdf = "N"
								Else
									gf.EliminaFileFisico(Server.MapPath(".") & "\" & pathPdf)
									pathPdf = "S"
								End If

								Ritorno &= pathPdf & ";"
								Ritorno &= QuanteFirmeValidate1 & ";"
								Ritorno &= QuanteFirmeValidate2 & ";"
								Ritorno &= QuanteFirmeValidate3 & ";"
								Ritorno &= QuanteFirmeValidate4 & ";"

								Rec.Close()


								Sql = "select * from giocatoritesserenfc where idGiocatore=" & idGiocatore & " And codsquadra='" & Squadra & "'"
								Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGen)
								Dim CodiceTessera As String = ""
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2
								Else
									If Rec2.Eof = False Then
										CodiceTessera = Rec2("codicetessera").Value
										Rec2.Close()

										Dim CashbackTotale As Single = 0
										Dim CashbackUtilizzato As Single = 0

										Sql = "select coalesce(sum(importo),0) As Importo from tesserenfc where numerotessera='" & CodiceTessera & "'"
										Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGen)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											CashbackTotale = Rec2("Importo").Value

											Rec2.Close
										End If

										Sql = "select coalesce(sum(importo),0) As Importo from cashbackutilizzato where codicetessera='" & CodiceTessera & "'"
										Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGen)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											CashbackUtilizzato = Rec2("Importo").Value

											Rec2.Close
										End If

										Dim CashbackDisponibile As Single = CashbackTotale - CashbackUtilizzato

										Ritorno &= CashbackTotale & ";"
										Ritorno &= CashbackDisponibile & ";"
										Ritorno &= CashbackUtilizzato & ";"
										Ritorno &= CodiceTessera & ";"

										Dim Movimenti As String = ""

										Sql = "select * From (" &
											"SELECT numerotessera, progressivo, descrizione, importo, dataora FROM tesserenfc " &
											"Where numerotessera = '" & CodiceTessera & "' " &
											"Union All " &
											"Select codicetessera as numerotessera, progressivo, 'Utilizzo Cashback' As descrizione, importo, dataesecuzione as dataora from cashbackutilizzato " &
											"Where codicetessera = '" & CodiceTessera & "' " &
											") as a " &
											"Order by dataora, progressivo"
										Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGen)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											Dim Pro As Integer = 1

											Do Until Rec2.Eof
												Dim DataOra As String = Rec2("DataOra").Value
												Dim d() As String = DataOra.Split(" ")
												DataOra = d(0)
												DataOra = Mid(DataOra, 7, 2) & "/" & Mid(DataOra, 5, 2) & "/" & Mid(DataOra, 1, 4) & " " & d(1)

												Movimenti &= Pro & "*" & Rec2("Descrizione").Value.replace("*", "_").replace(";", "_") & "*" & Rec2("Importo").Value & "*" & dataora & "§"
												Pro += 1

												Rec2.MoveNext
											Loop

											Rec2.Close
										End If

										Ritorno &= Movimenti & ";"
									Else
										Ritorno &= "0.00;"
										Ritorno &= "0.00;"
										Ritorno &= "0.00;"
										Ritorno &= ";"
										Ritorno &= ";"
									End If
								End If
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ": " & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Try
					If TipoDB = "SQLSERVER" Then
						Sql = "SELECT IsNull(Max(idGiocatore),0)+1 FROM Giocatori"
					Else
						Sql = "SELECT Coalesce(Max(idGiocatore),0)+1 FROM Giocatori"
					End If
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						'If Rec(0).Value Is DBNull.Value Then
						'	idGioc = 1
						'Else
						idGioc = Rec(0).Value
						'End If
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
								   idTaglia As String, Modalita As String, Cap As String, CittaNascita As String, Mittente As String) As String
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
				Dim idGioc As Integer = -1
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Dim Maggiorenne As String = "N"

					'If idGiocatore = "-1" Then
					If Modalita = "INSERIMENTO" Then
						Sql = "SELECT * FROM Giocatori Where idAnno=" & idAnno & " And Upper(lTrim(rTrim(CodFiscale)))='" & CodFiscale.ToUpper.Trim & "' And Eliminato='N'"
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If Not Rec.Eof() Then
							Ritorno = StringaErrore & "Codice fiscale già presente in archivio" ' : " & CodFiscale & "--->" & Sql
							Ok = False
						Else
							Ritorno = ""
						End If
						Rec.Close()

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

							If TipoDB = "SQLSERVER" Then
								Sql = "Select IsNull(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
							Else
								Sql = "Select Coalesce(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
							End If
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								'If Rec(0).Value Is DBNull.Value Then
								'	idUtente = 1
								'Else
								idUtente = Rec(0).Value
								'End If
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
										"'N', " &
										"'" & stringaWidgets & "' " &
										")"
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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

								Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, Body, EMail, {""})
							End If
						End If
					Else
						Sql = "SELECT * FROM Giocatori Where idAnno=" & idAnno & " And Upper(lTrim(rTrim(CodFiscale)))='" & CodFiscale.ToUpper.Trim & "'"
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If Not Rec.Eof() Then
							Maggiorenne = "" & Rec("Maggiorenne").Value
							'Ok = False
						End If
						Rec.Close()
					End If

					'If Ok Then
					'	Try
					'		Sql = "SELECT Max(idGiocatore)+1 FROM Giocatori"
					'		Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
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
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If Rec.Eof() Then
							GiaCe = False
							' Dim conta As Integer = 0

							'Do While Ritorno.Contains(StringaErrore) Or Ritorno = ""
							'Try
							'	Sql = "Delete  From Giocatori Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
							'	Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
						Rec.Close()
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
									" " & idGiocatore & ", " &
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
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
								'Else
								'	Ritorno = CalcolaSemafori(Conn, Connessione, Squadra, idGiocatore)
								'	If Ritorno.Contains(StringaErrore) Then
								'		Ok = False
								'	End If
							End If
						End If
					End If
				Else
					Ok = False
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

					Dim ct As String = CreaNumeroTesseraNFC(Server.MapPath("."), Conn, Connessione, Squadra, idGiocatore)
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
	Public Function CreaNumeroTesseraNFCDaFuori(Squadra As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Ritorno = CreaNumeroTesseraNFC(Server.MapPath("."), Conn, Connessione, Squadra, idGiocatore)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Update Giocatori Set Eliminato='S' Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
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
	Public Function AggiungeCategoriaAGiocatore(Squadra As String, ByVal idAnno As String, idGiocatore As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Dim Giocatori() As String = idGiocatore.Split(";")

					For Each g As String In Giocatori
						If g <> "" Then
							Try
								If TipoDB = "SQLSERVER" Then
									Sql = "Update Giocatori Set Categorie = Categorie + '" & idCategoria & "-' Where idAnno=" & idAnno & " And idGiocatore=" & g
								Else
									Sql = "Update Giocatori Set Categorie = Concat(Categorie, '" & idCategoria & "-') Where idAnno=" & idAnno & " And idGiocatore=" & g
								End If
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
	Public Function EliminaGiocatoreDallaCategoria(Squadra As String, ByVal idAnno As String, idGiocatore As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Update Giocatori Set Categorie = Replace(Categorie, '" & idCategoria & "-', '') Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

						Sql = "Update Giocatori Set idCategoria=-1 Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And idCategoria=" & idCategoria
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

						If Ok Then
							Sql = "Update Giocatori Set idCategoria2=-1 Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And idCategoria2=" & idCategoria
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If
						End If

						If Ok Then
							Sql = "Update Giocatori Set idCategoria3=-1 Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And idCategoria3=" & idCategoria
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
	Public Function SalvaPagamento(Squadra As String, idAnno As String, idGiocatore As String, Pagamento As String, Commento As String,
								   idPagatore As String, idRegistratore As String, Note As String, Validato As String, idTipoPagamento As String,
								   idRata As String, idQuota As String, Suffisso As String, sNumeroRicevuta As String, DataRicevuta As String, idUtente As String,
								   idModalitaPagamento As String, ImportoManuale As String, DescrizioneManuale As String, DataManuale As String) As String
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

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If sNumeroRicevuta <> "" And sNumeroRicevuta <> "Bozza" Then
					Sql = "SELECT * FROM GiocatoriPagamenti Where NumeroRicevuta='" & sNumeroRicevuta & "'"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If Not Rec.Eof() Then
						Ritorno = StringaErrore & " Numero ricevuta già presente"
						Ok = False
					End If
					Rec.Close()
				End If

				If Not Ritorno.Contains(StringaErrore) Then
					Dim Progressivo As Integer
					Dim ProgressivoGenerale As Integer

					'Dim DataPagamento As String = Now.Year & "-" & Format(Now.Month, "00") & "-" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
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
					Dim NumeroRicevuta As String = ""

					Try
						Sql = "SELECT * FROM Anni"
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
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
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
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
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
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
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
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
								If sNumeroRicevuta <> "" Then
									NumeroRicevuta = sNumeroRicevuta
								Else
									If Validato = "S" Then
										'If TipoDB = "SQLSERVER" Then
										'	Sql = "SELECT IsNull(Max(Progressivo),0)+1 As Progressivo FROM DatiFattura Where Anno=" & Now.Year
										'Else
										'	Sql = "SELECT Coalesce(Max(Progressivo),0)+1 As Progressivo FROM DatiFattura Where Anno=" & Now.Year
										'End If
										'Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
										''If Rec(0).Value Is DBNull.Value Then
										''	ProgressivoGenerale = 1
										''	Sql = "Insert Into DatiFattura Values(" & Now.Year & ", 1)"
										''Else
										'ProgressivoGenerale = Rec("Progressivo").Value
										Rec.Close

										Sql = "Select * From DatiFattura Where Anno=" & Now.Year
										Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
										If Rec.Eof() Then
											ProgressivoGenerale = 1
											Sql = "Insert Into datifattura Values(" & Now.Year & ", 1)"
										Else
											ProgressivoGenerale = Rec("Progressivo").Value + 1
											Sql = "Update datifattura Set Progressivo = " & ProgressivoGenerale & " Where Anno=" & Now.Year
										End If
										Rec.Close

										'If ProgressivoGenerale = 0 Then
										'	Sql = "Insert Into datifattura Values(" & Now.Year & ", 1)"
										'Else
										'	Sql = "Update datifattura Set Progressivo = " & ProgressivoGenerale + 1 & " Where Anno=" & Now.Year
										'End If
										''End If
										'Rec.Close()

										If Suffisso <> "" Then
											NumeroRicevuta = ProgressivoGenerale & "/" & Suffisso & "/" & Now.Year
										Else
											NumeroRicevuta = ProgressivoGenerale & "/" & Now.Year
										End If

										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If
									Else
										NumeroRicevuta = "Bozza"
									End If
								End If

								If Ok Then
									If TipoDB = "SQLSERVER" Then
										Sql = "SELECT IsNull(Max(Progressivo),0)+1 FROM GiocatoriPagamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
									Else
										Sql = "SELECT Coalesce(Max(Progressivo),0)+1 FROM GiocatoriPagamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
									End If
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									'If Rec(0).Value Is DBNull.Value Then
									'	Progressivo = 1
									'Else
									Progressivo = Rec(0).Value
									'End If
									Rec.Close()

									Sql = "Insert Into GiocatoriPagamenti Values (" &
										" " & idAnno & ", " &
										" " & idGiocatore & ", " &
										" " & Progressivo & ", " &
										" " & Pagamento & ", " &
										"'" & DataRicevuta & "', " &
										"'N', " &
										"'" & Commento.Replace("'", "''") & "', " &
										" " & idPagatore & ", " &
										" " & idRegistratore & ", " &
										"'" & Note.Replace("'", "''") & "', " &
										"'" & Validato & "', " &
										" " & idTipoPagamento & ", " &
										"'" & idRata & "', " &
										" " & idQuota & ", " &
										"'" & NumeroRicevuta & "', " &
										" " & idModalitaPagamento & ", " &
										" " & ImportoManuale & ", " &
										"'" & DescrizioneManuale.Replace("'", "''") & "', " &
										"'" & DataManuale & "' " &
										")"
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									Else
										Ritorno = Progressivo
									End If
								End If
							End If
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					'If Ok And NumeroRicevuta <> "Bozza" Then
					'	Try
					'		Dim gf As New GestioneFilesDirectory
					'		Dim filePaths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
					'		Dim p() As String = filePaths.Split(";")
					'		If Strings.Right(p(0), 1) <> "\" Then
					'			p(0) &= "\"
					'		End If
					'		p(2) = p(2).Replace(vbCrLf, "").Trim
					'		If Strings.Right(p(2), 1) <> "/" Then
					'			p(2) = p(2) & "/"
					'		End If
					'		' Dim url As String = p(2) & NomeSquadra.Replace(" ", "_") & "/Societa/" & idAnno & "_1.jpg"

					'		Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					'		pp = pp.Replace(vbCrLf, "").Trim
					'		If Strings.Right(pp, 1) = "\" Then
					'			pp = Mid(pp, 1, pp.Length - 1)
					'		End If
					'		Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

					'		Dim nomeImm As String = p(2) & NomeSquadra.Replace(" ", "_") & "/Societa/" & idAnno & "_1.kgb"
					'		Dim pathImm As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\" & idAnno & "_1.kgb"
					'		Dim nomeImmConv As String = ""
					'		Dim c As New CriptaFiles
					'		If ControllaEsistenzaFile(pathImm) Then
					'			nomeImmConv = p(2) & "Appoggio/Societa_" & idAnno & "_1_" & Esten & ".png"
					'			Dim pathImmConv As String = pp & "\Appoggio\Societa_" & idAnno & "_1_" & Esten & ".png"
					'			c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)
					'		End If

					'		Dim pathRicevuta As String = p(0) & Squadra & "\Scheletri\ricevuta_pagamento.txt"
					'		If Not ControllaEsistenzaFile(pathRicevuta) Then
					'			pathRicevuta = Server.MapPath(".") & "\Scheletri\ricevuta_pagamento.txt"
					'		End If
					'		Dim Body As String = gf.LeggeFileIntero(pathRicevuta)
					'		Dim path As String = p(0) & "\" & Squadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\"
					'		gf.CreaDirectoryDaPercorso(path)
					'		Dim fileFinale As String = path & "Ricevuta_" & Progressivo & ".pdf"
					'		Dim fileAppoggio As String = path & "Ricevuta_" & Progressivo & ".html"


					'		Body = Body.Replace("***URL LOGO***", nomeImmConv)
					'		Body = Body.Replace("***NOME POLISPORTIVA***", NomePolisportiva)
					'		Body = Body.Replace("***INDIRIZZO***", Indirizzo)
					'		Body = Body.Replace("***DATI***", Dati)
					'		Body = Body.Replace("***NUMERO_RICEVUTA***", NumeroRicevuta)
					'		Dim d() As String = DataRicevuta.Split("-")
					'		Dim sDataRicevuta As String = d(2) & "/" & d(1) & "/" & d(0)
					'		Body = Body.Replace("***DATA_RICEVUTA***", sDataRicevuta) ' Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
					'		Body = Body.Replace("***NOME***", Cognome & " " & Nome)
					'		Body = Body.Replace("***MOTIVAZIONE***", CognomeIscritto & " " & NomeIscritto & " " & CodFiscaleIscritto & " " & Altro)
					'		Body = Body.Replace("***IMPORTO***", Intero)
					'		Body = Body.Replace("***VIRGOLE***", Virgola)

					'		Dim Cifre1 As String = convertNumberToReadableString(Val(Intero))
					'		Dim Cifre2 As String = convertNumberToReadableString(Val(Virgola))
					'		Dim Altro2 As String = ""
					'		If Cifre2 <> "" Then
					'			Altro2 = "/" & Virgola
					'		End If
					'		Body = Body.Replace("***IMPORTO LETTERE***", Cifre1 & Altro2)

					'		filePaths = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					'		filePaths = filePaths.Replace(vbCrLf, "").Trim
					'		If Strings.Right(filePaths, 1) <> "\" Then
					'			filePaths &= "\"
					'		End If
					'		' Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & idGiocatore & "_" & idPagatore & ".png"
					'		' Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Segreteria\" & idAnno & ".kgb"

					'		Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_").Trim & "\Utenti\" & idAnno & "_" & idUtente & "_Firma.kgb"
					'		' Return pathFirma
					'		If ControllaEsistenzaFile(pathFirma) Then
					'			Dim urlFirma As String = pp & "\" & NomeSquadra.Replace(" ", "_").Trim & "\Utenti\" & idAnno & "_" & idUtente & "_Firma.kgb"
					'			'Dim pathFirmaConv As String = p(2) & "/Appoggio/Firma_" & Esten & ".png"
					'			Dim urlFirmaConv As String = pp & "\Appoggio\Firma_" & Esten & ".png"
					'			c.DecryptFile(CryptPasswordString, urlFirma, urlFirmaConv)

					'			Body = Body.Replace("***URL FIRMA***", urlFirmaConv)
					'		Else
					'			Body = Body.Replace("***URL FIRMA***", "")
					'		End If

					'		gf.EliminaFileFisico(fileAppoggio)
					'		gf.ApreFileDiTestoPerScrittura(fileAppoggio)
					'		gf.ScriveTestoSuFileAperto(Body)

					'		gf.ChiudeFileDiTestoDopoScrittura()

					'		' Scontrino
					'		Dim pathScontr As String = p(0) & Squadra & "\Scheletri\ricevuta_scontrino.txt"
					'		If Not ControllaEsistenzaFile(pathScontr) Then
					'			pathScontr = Server.MapPath(".") & "\Scheletri\ricevuta_scontrino.txt"
					'		End If
					'		Dim BodyScontrino As String = gf.LeggeFileIntero(pathScontr)
					'		Dim pathScontrino As String = p(0) & "\" & Squadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\"
					'		gf.CreaDirectoryDaPercorso(pathScontrino)
					'		Dim fileFinaleScontrino As String = path & "Scontrino_" & idPagamento & ".pdf"
					'		Dim fileAppoggioScontrino As String = path & "Scontrino_" & idPagamento & ".html"
					'		BodyScontrino = BodyScontrino.Replace("***NOME POLISPORTIVA***", NomePolisportiva)
					'		BodyScontrino = BodyScontrino.Replace("***INDIRIZZO***", Indirizzo)
					'		BodyScontrino = BodyScontrino.Replace("***DATI***", Dati)
					'		If NumeroRicevuta <> "" Then
					'			BodyScontrino = BodyScontrino.Replace("***NUMERO_RICEVUTA***", NumeroRicevuta)
					'		Else
					'			If Suffisso <> "" Then
					'				BodyScontrino = BodyScontrino.Replace("***NUMERO_RICEVUTA***", idPagamento & "/" & Suffisso & "/" & Now.Year)
					'			Else
					'				BodyScontrino = BodyScontrino.Replace("***NUMERO_RICEVUTA***", idPagamento & "/" & Now.Year)
					'			End If
					'		End If
					'		If DataRicevuta <> "" Then
					'			Dim d() As String = DataRicevuta.Split("-")
					'			Dim sDataRicevuta As String = d(2) & "/" & d(1) & "/" & d(0)
					'			BodyScontrino = BodyScontrino.Replace("***DATA_RICEVUTA***", sDataRicevuta) ' Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
					'		Else
					'			BodyScontrino = BodyScontrino.Replace("***DATA_RICEVUTA***", Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
					'		End If
					'		BodyScontrino = BodyScontrino.Replace("***MOTIVAZIONE***", CognomeIscritto & " " & NomeIscritto & " " & CodFiscaleIscritto & " " & Altro)
					'		BodyScontrino = BodyScontrino.Replace("***IMPORTO***", Intero)

					'		gf.EliminaFileFisico(fileAppoggioScontrino)
					'		gf.ApreFileDiTestoPerScrittura(fileAppoggioScontrino)
					'		gf.ScriveTestoSuFileAperto(BodyScontrino)
					'		gf.ChiudeFileDiTestoDopoScrittura()
					'		' Scontrino

					'		Dim pp2 As New pdfGest
					'		Ritorno = pp2.ConverteHTMLInPDF(fileAppoggio, fileFinale, "")
					'		Dim Ritorno2 As String = pp2.ConverteHTMLInPDF(fileAppoggioScontrino, fileFinaleScontrino, "")
					'		If Ritorno<> "OK" And Ritorno2<> "OK" Then
					'			Ok = False
					'		Else
					'			If Ritorno2<> "OK" Then
					'				Ritorno = Ritorno2
					'			End If
					'		End If
					'	Catch ex As Exception
					'		Ritorno = StringaErrore & " " & ex.Message
					'	End Try
					'End If
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
	Public Function ModificaPagamento(Squadra As String, idPagamento As String, idAnno As String, idGiocatore As String, Pagamento As String, Commento As String,
								   idPagatore As String, idRegistratore As String, Note As String, Validato As String, idTipoPagamento As String,
								   idRata As String, idQuota As String, Suffisso As String, NumeroRicevuta As String, DataRicevuta As String, idUtente As String,
								   idModalitaPagamento As String, Stato As String, Modifica As String, ImportoManuale As String, DescrizioneManuale As String, DataManuale As String,
									  idOperatore As String) As String
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

				Dim Cognome As String = ""
				Dim CognomePagatore As String = ""
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
				Dim indirizzoPagatore As String = ""
				Dim nuovoIdPagamento As Integer

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "-----------------------------------")
				If Not Ritorno.Contains(StringaErrore) Then
					ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Modifica: " & Modifica)
					ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "ID Pagamento: " & idPagamento)
					ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Numero ricevuta: " & NumeroRicevuta)
					ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Validato: " & Validato)

					If Modifica <> "AMM-SUPERUSER" And Stato <> "Bozza" Then
						If NumeroRicevuta <> "" And NumeroRicevuta <> "Bozza" And Validato = "N" Then
							Sql = "SELECT * FROM GiocatoriPagamenti Where NumeroRicevuta='" & NumeroRicevuta & "'"
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If Not Rec.Eof() Then
								Ritorno = StringaErrore & " Numero ricevuta già presente"
								ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
								Ok = False
							End If
							Rec.Close()
						End If
					End If

					If Ok Then
						Try
							Sql = "SELECT * FROM Anni"
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If Rec.Eof() Then
								Ritorno = StringaErrore & " Nessuna squadra rilevata"
								ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
								Ok = False
							Else
								NomeSquadra = "" & Rec("NomeSquadra").Value
								NomePolisportiva = "" & Rec("NomePolisportiva").Value
								Indirizzo = "" & Rec("Indirizzo").Value
								CodiceFiscale = "" & Rec("CodiceFiscale").Value
								PIva = "" & Rec("PIva").Value
								Telefono = "" & Rec("Telefono").Value
								eMail = "" & Rec("Mail").Value
							End If
							'Rec.Close()
							ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Nome squadra: " & NomeSquadra)

							If Ok Then
								ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "ID Pagatore: " & idPagatore)
								If idPagatore = 3 Then
									Sql = "SELECT * FROM Giocatori Where idGiocatore=" & idGiocatore
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If Rec.Eof() Then
										Ritorno = StringaErrore & " Nessun giocatore rilevato"
										ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
										Ok = False
									Else
										Cognome = "" & Rec("Cognome").Value
										Nome = "" & Rec("Nome").Value
										CodFiscalePagatore = "" & Rec("CodFiscale").Value

										CognomeIscritto = "" & Rec("Cognome").Value
										NomeIscritto = "" & Rec("Nome").Value
										CodFiscaleIscritto = "" & Rec("CodFiscale").Value
									End If
									'Rec.Close()
								Else
									Sql = "SELECT * FROM Giocatori Where idGiocatore=" & idGiocatore
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If Rec.Eof() Then
										Ritorno = StringaErrore & " Nessun giocatore rilevato"
										ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
										Ok = False
									Else
										CognomeIscritto = "" & Rec("Cognome").Value
										NomeIscritto = "" & Rec("Nome").Value
										CodFiscaleIscritto = "" & Rec("CodFiscale").Value
									End If
									'Rec.Close()

									Sql = "SELECT * FROM GiocatoriDettaglio Where idGiocatore=" & idGiocatore
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If Rec.Eof() Then
										Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
										ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
										Ok = False
									Else
										If idPagatore = 1 Then
											CognomePagatore = "" & Rec("Genitore1").Value
											indirizzoPagatore = "" & Rec("Indirizzo1").Value
											CodFiscalePagatore = "" & Rec("CodFiscale1").Value
										Else
											CognomePagatore = "" & Rec("Genitore2").Value
											indirizzoPagatore = "" & Rec("Indirizzo2").Value
											CodFiscalePagatore = "" & Rec("CodFiscale2").Value
										End If
									End If
									'Rec.Close()
								End If
								ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Cognome: " & Cognome)
								ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Cognome Iscritto: " & CognomeIscritto)
								ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Cognome Pagatore: " & CognomePagatore)

								If Ok Then
									Dim Altro As String = ""
									Dim ProgressivoGenerale As Integer

									If NumeroRicevuta = "Bozza" Then
										If Validato = "S" Then
											'If TipoDB = "SQLSERVER" Then
											'	Sql = "SELECT IsNull(Max(Progressivo),0)+1 FROM DatiFattura Where Anno=" & Now.Year
											'Else
											'	Sql = "SELECT Coalesce(Max(Progressivo),0)+1 FROM DatiFattura Where Anno=" & Now.Year
											'End If
											'Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
											'' If Rec(0).Value Is DBNull.Value Then
											''If Rec(0).Value = 1 Then
											''	ProgressivoGenerale = 1
											''	Sql = "Insert Into DatiFattura Values(" & Now.Year & ", 1)"
											''Else
											'ProgressivoGenerale = Rec(0).Value
											'Sql = "Update DatiFattura Set Progressivo = " & ProgressivoGenerale & " Where Anno=" & Now.Year
											''End If
											'Rec.Close()

											Sql = "Select * From DatiFattura Where Anno=" & Now.Year
											Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
											If Rec.Eof() Then
												ProgressivoGenerale = 1
												Sql = "Insert Into datifattura Values(" & Now.Year & ", 1)"
												ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Nuova riga in dati fattura: 1")
											Else
												ProgressivoGenerale = Rec("Progressivo").Value + 1
												Sql = "Update datifattura Set Progressivo = " & ProgressivoGenerale & " Where Anno=" & Now.Year
												ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Modifica riga in dati fattura: " & ProgressivoGenerale)
											End If
											Rec.Close

											Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
											If Ritorno.Contains(StringaErrore) Then
												ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
												Ok = False
											End If

											If Suffisso <> "" Then
												NumeroRicevuta = ProgressivoGenerale & "/" & Suffisso & "/" & Now.Year
											Else
												NumeroRicevuta = ProgressivoGenerale & "/" & Now.Year
											End If
											ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Nuovo numero ricevuta: " & NumeroRicevuta)
										Else
											NumeroRicevuta = "Bozza"
											ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Nuovo numero ricevuta: " & NumeroRicevuta)
										End If
									End If

									If NumeroRicevuta <> "" And NumeroRicevuta <> "Bozza" Then
										Altro = ", NumeroRicevuta = '" & NumeroRicevuta & "' "
									End If

									'Sql = "Delete From GiocatoriPagamenti Where idGiocatore = " & idGiocatore & " And Progressivo = " & idPagamento
									'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									'If Ritorno.Contains(StringaErrore) Then
									'	Ok = False
									'End If

									'If Ok Then
									'	Sql = "Insert Into GiocatoriPagamenti Values (" &
									'		" " & idAnno & ", " &
									'		" " & idGiocatore & ", " &
									'		" " & idPagamento & ", " &
									'		" " & Pagamento & ", " &
									'		"'" & DataRicevuta & "', " &
									'		"'" & Validato & "', " &
									'		"'" & Commento.Replace("'", "''") & "', " &
									'		" " & idPagatore & ", " &
									'		" " & idRegistratore & ", " &
									'		"'" & Note.Replace("'", "''") & "', " &
									'		"'" & Validato & "', " &
									'		" " & idTipoPagamento & ", " &
									'		"'" & idRata & "', " &
									'		" " & idQuota & ", " &
									'		"'" & NumeroRicevuta & "', " &
									'		" " & idModalitaPagamento & " " &
									'		")"

									If TipoDB = "SQLSERVER" Then
										Sql = "SELECT IsNull(Max(Progressivo),0) + 1 FROM GiocatoriPagamenti Where idGiocatore=" & idGiocatore
									Else
										Sql = "SELECT Coalesce(Max(Progressivo),0) + 1 FROM GiocatoriPagamenti Where idGiocatore=" & idGiocatore
									End If
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									'If Rec(0).Value Is DBNull.Value Then
									'	nuovoIdPagamento = 1
									'Else
									nuovoIdPagamento = Rec(0).Value
									'End If
									Rec.Close()

									ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "Nuovo id pagamento: " & nuovoIdPagamento)

									Sql = "Update GiocatoriPagamenti Set " &
										"Progressivo=" & nuovoIdPagamento & ", " &
										"Pagamento=" & Pagamento & ", " &
										"DataPagamento='" & DataRicevuta & "', " &
										"Commento='" & Commento.Replace("'", "''") & "', " &
										"idUtentePagatore=" & idPagatore & ", " &
										"idUtenteRegistratore=" & idRegistratore & ", " &
										"Note='" & Note.Replace("'", "''") & "', " &
										"Validato='" & Validato & "', " &
										"idTipoPagamento=" & idTipoPagamento & ", " &
										"idRata='" & idRata & "', " &
										"idQuota=" & idQuota & ", " &
										"MetodoPagamento=" & idModalitaPagamento & ", " &
										"ImportoManuale=" & ImportoManuale & ", " &
										"DescrizioneManuale='" & DescrizioneManuale.Replace("'", "''") & "', " &
										"DataManuale='" & DataManuale & "' " &
										Altro &
										"Where idGiocatore = " & idGiocatore & " And Progressivo = " & idPagamento
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
										Ok = False
									End If
								End If
								'End If
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", Ritorno)
							Ok = False
						End Try
					End If

					If Ok Then
						Sql = "commit"
						Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

						ScriveLog(Server.MapPath("."), Squadra, "Pagamenti", "---------------------------------------------------")
						ScriveLog(Server.MapPath("."), Squadra, "Pagamenti", "Genera Ricevute e scontrini per numero ricevuta " & NumeroRicevuta)
						If NumeroRicevuta <> "Bozza" Then
							Ritorno = GeneraRicevutaEScontrino(Server.MapPath("."), Squadra, NomeSquadra, idAnno, idGiocatore, nuovoIdPagamento, idUtente, nuovoIdPagamento, idOperatore)
							If Ritorno = "*" Or Ritorno = "OK" Then
								Ritorno = nuovoIdPagamento
							End If
							ScriveLog(Server.MapPath("."), Squadra, "Pagamenti", "---------------------------------------------------")
						Else
							If Ritorno = "*" Or Ritorno = "OK" Then
								Ritorno = nuovoIdPagamento
							End If
						End If
						ScriveLog(Server.MapPath("."), Squadra, "Pagamenti", "Ritorno genera ricevuta: " & Ritorno)
					End If
				End If

				If Not Ok Then
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "ROLLBACK")
				End If

				ScriveLog(Server.MapPath("."), Squadra, "ModificaPagamenti", "-----------------------------------")
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RistampaRicevutaScontrino(Squadra As String, NomeSquadra As String, idAnno As String, idGiocatore As String, idPagamento As String, idUtente As String, idOperatore As String) As String
		Dim Ritorno As String = ""

		Ritorno = GeneraRicevutaEScontrino(Server.MapPath("."), Squadra, NomeSquadra, idAnno, idGiocatore, idPagamento, idUtente, idPagamento, idOperatore)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaPagamentoGiocatore(Squadra As String, idAnno As String, idGiocatore As String, Progressivo As String, idRegistratore As String) As String
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

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Update GiocatoriPagamenti Set " &
							"Eliminato='S', idQuota = -1, idRata = '' " &
							"Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And Progressivo=" & Progressivo
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
							Dim ora As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
							Sql = "Insert Into GiocatoriPagamentiEliminazioni Values (" & idAnno & ", " & idGiocatore & ", " & Progressivo & ", " & idRegistratore & ", '" & ora & "')"
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
	Public Function RitornaPagamentiGiocatore(Squadra As String, idAnno As String, idGiocatore As String) As String
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
				Dim TotPag As Single = 0

				If TipoDB = "SQLSERVER" Then
					Sql = "Select IsNull(TotalePagamento,0) As TotalePagamento From GiocatoriDettaglio Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
				Else
					Sql = "Select Coalesce(TotalePagamento,0) As TotalePagamento From GiocatoriDettaglio Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
				End If
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						'If Not Rec("TotalePagamento").Value Is DBNull.Value Then
						TotPag = Rec("TotalePagamento").Value
						'Else
						'	TotPag = 0
						'End If
						Rec.Close()

						Sql = "Select * From GiocatoriPagamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And Eliminato='N' And NumeroRicevuta<>'Bozza' Order By Progressivo"
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim Totale As Single = 0
							'Ritorno = "Totale a pagare;" & Format(TotPag, "#0.#0") & ";;§"
							Dim Ritorno2 As String = ""
							Do Until Rec.Eof()
								Dim desc As String = "Rata Quota"
								If Rec("idTipoPagamento").Value = 2 Then
									desc = "Altro"
								End If

								Ritorno2 &= Rec("Progressivo").Value & ";" & Rec("Pagamento").Value & ";" & Rec("DataPagamento").Value & ";" & Rec("Commento").Value & ";" & Rec("ImportoManuale").Value & ";" & Rec("idTipoPagamento").Value & ";" & desc & "§"
								Totale += (Rec("Pagamento").Value)

								Rec.MoveNext()
							Loop
							Rec.Close()

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
		'Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		'If Connessione = "" Then
		'	Ritorno = ErroreConnessioneNonValida
		'Else
		'	Dim Conn As Object = New clsGestioneDB(Squadra)

		'	If TypeOf (Conn) Is String Then
		'		Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
		'	Else
		'		Dim Rec as object
		'		Dim Sql As String

		'		Sql = "Select NomeSquadra, Descrizione, iscrFirmaEntrambi From Anni Where idAnno = " & Anno
		'		Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
		'		If TypeOf (Rec) Is String Then
		'			Ritorno = Rec
		'		Else
		'			If Rec.Eof() Then
		'				Ritorno = StringaErrore & " Nessuna squadra rilevata"
		'			Else
		'				Dim NomeSquadra As String = "" & Rec("NomeSquadra").Value
		'				Dim Descrizione As String = "" & Rec("Descrizione").Value
		'				Dim iscrFirmaEntrambi As String = "" & Rec("iscrFirmaEntrambi").Value
		'				Rec.Close()

		'				Sql = "Select MailGenitore1, MailGenitore2, B.Cognome + ' ' + B.Nome As Genitore3 , Genitore1, Genitore2, MailGenitore3, " &
		'								"B.Maggiorenne, GenitoriSeparati, AffidamentoCongiunto, idTutore " &
		'								"From GiocatoriDettaglio A " &
		'								"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
		'								"Where A.idGiocatore = " & idGiocatore
		'				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
		'				If TypeOf (Rec) Is String Then
		'					Ritorno = Rec
		'				Else
		'					If Rec.Eof() Then
		'						Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
		'					Else
		'						Dim Maggiorenne As String = "" & Rec("Maggiorenne").Value
		'						Dim GenitoriSeparati As String = "" & Rec("GenitoriSeparati").Value
		'						Dim AffidamentoCongiunto As String = "" & Rec("AffidamentoCongiunto").Value
		'						Dim idTutore As String = "" & Rec("idTutore").Value
		'						Dim ceGenitore1 As String = "" & Rec("Genitore1").Value
		'						Dim ceGenitore2 As String = "" & Rec("Genitore2").Value
		'						Rec.Close()

		'Dim gf As New GestioneFilesDirectory
		'Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		'Dim P() As String = PathAllegati.Split(";")
		'If Strings.Right(P(0), 1) = "\" Then
		'	P(0) = Mid(P(0), 1, P(0).Length - 1)
		'End If
		'Dim fileDaCopiare As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".html"
		'Dim fileDaCopiarePDF As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".pdf"
		'Dim fileLog As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".log"

		''Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
		'gf.CreaDirectoryDaPercorso(fileDaCopiare)
		'gf.EliminaFileFisico(fileDaCopiare)
		'gf.EliminaFileFisico(fileDaCopiarePDF)
		'gf.EliminaFileFisico(fileLog)
		' Dim fileScheletro As String = Server.MapPath(".") & "\Scheletri\base_iscrizione_.txt"

		'Dim fileScheletro As String = P(0) & "\" & Squadra & "\Scheletri\base_iscrizione_.txt"
		'If Not ControllaEsistenzaFile(fileScheletro) Then
		'	fileScheletro = HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_iscrizione_.txt"
		'End If

		'If ControllaEsistenzaFile(fileScheletro) Then
		Try
			'Dim fileFirme As String = gf.LeggeFileIntero(fileScheletro)
			'fileFirme = RiempieFileFirme(fileFirme, Anno, idGiocatore, Rec, Conn, Connessione, NomeSquadra, P, Descrizione)
			Dim gt As New GestioneTags(Server.MapPath("."))
			Dim fileFirme As String = gt.EsegueFileFirme(Server.MapPath("."), Squadra, idGiocatore, Anno, False)
			'Return fileFirme

			'gf.EliminaFileFisico(fileDaCopiare)
			'gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
			'gf.ScriveTestoSuFileAperto(fileFirme)
			'gf.ChiudeFileDiTestoDopoScrittura()

			''File.Copy(fileDaCopiare, fileDaCopiare2)
			'Dim pp As New pdfGest
			'Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

			'If Ritorno = "*" Then
			'	gf.EliminaFileFisico(fileDaCopiare)
			'End If

			'gf = Nothing
		Catch ex As Exception
			Ritorno = StringaErrore & " " & ex.Message
		End Try
		'Else
		'	Ritorno = StringaErrore & " Scheletro iscrizione non trovato"
		'End If
		'gf = Nothing
		'					End If
		'				End If
		'			End If
		'		End If
		'	End If
		'End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ControllaEsistenzaModuloAssociato(Squadra As String, Anno As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		'Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		'If Connessione = "" Then
		'	Ritorno = ErroreConnessioneNonValida
		'Else
		'	Dim Conn As Object = New clsGestioneDB(Squadra)

		'	If TypeOf (Conn) Is String Then
		'		Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
		'	Else
		'		Dim Rec as object
		'		Dim Sql As String

		'Sql = "Select NomeSquadra, Descrizione, iscrFirmaEntrambi From Anni Where idAnno = " & Anno
		'Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
		'If TypeOf (Rec) Is String Then
		'	Ritorno = Rec
		'Else
		'	If Rec.Eof() Then
		'		Ritorno = StringaErrore & " Nessuna squadra rilevata"
		'	Else
		'		Dim NomeSquadra As String = Rec("NomeSquadra").Value
		'		Dim AnnoTabella As String = Rec("Descrizione").Value
		'		Dim iscrFirmaEntrambi As String = "" & Rec("iscrFirmaEntrambi").Value
		'		Rec.Close()

		'Dim gf As New GestioneFilesDirectory
		'Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		'Dim P() As String = PathAllegati.Split(";")
		'If Strings.Right(P(0), 1) = "\" Then
		'	P(0) = Mid(P(0), 1, P(0).Length - 1)
		'End If
		'Dim fileDaCopiare As String = P(0) & "\" & Squadra & "\Firme\associato_" & Anno & "_" & idGiocatore & ".html"
		'Dim fileDaCopiarePDF As String = P(0) & "\" & Squadra & "\Firme\associato_" & Anno & "_" & idGiocatore & ".pdf"
		'Dim fileLog As String = P(0) & "\" & Squadra & "\Firme\associato_" & Anno & "_" & idGiocatore & ".log"

		''Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
		'gf.CreaDirectoryDaPercorso(fileDaCopiare)
		'gf.EliminaFileFisico(fileDaCopiare)
		'gf.EliminaFileFisico(fileDaCopiarePDF)
		'gf.EliminaFileFisico(fileLog)

		Dim gT As New GestioneTags(Server.MapPath("."))
		Ritorno = gT.EsegueFileAssociato(Server.MapPath("."), Squadra, idGiocatore, Anno)

		''File.Copy(fileDaCopiare, fileDaCopiare2)
		'Dim pp As New pdfGest
		'Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

		'If Ritorno = "*" Then
		'	gf.EliminaFileFisico(fileDaCopiare)
		'End If

		'		Sql = "Select MailGenitore1, MailGenitore2, B.Cognome + ' ' + B.Nome As Genitore3 , Genitore1, Genitore2, MailGenitore3, " &
		'				"B.Maggiorenne, GenitoriSeparati, AffidamentoCongiunto, idTutore, A.DataDiNascita1, A.DataDiNascita2, B.DataDiNascita As DataDiNascita3, " &
		'				"A.CittaNascita1, A.CittaNascita2, B.CittaNascita As CittaNascita3, " &
		'				"A.Citta1, A.Citta2, B.Citta As Citta3, " &
		'				"A.Indirizzo1, A.Indirizzo2, B.Indirizzo As Indirizzo3, " &
		'				"A.TelefonoGenitore1, A.TelefonoGenitore2, B.Telefono As Telefono3, " &
		'				"A.MailGenitore1, A.MailGenitore2, B.EMail As EMail3, B.Cognome + ' ' + B.Nome As Giocatore " &
		'				"From GiocatoriDettaglio A " &
		'				"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
		'				"Where A.idGiocatore = " & idGiocatore
		'Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
		'If TypeOf (Rec) Is String Then
		'	Ritorno = Rec
		'Else
		'	If Rec.Eof() Then
		'		Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
		'	Else
		'		Dim Maggiorenne As String = "" & Rec("Maggiorenne").Value
		'		Dim GenitoriSeparati As String = "" & Rec("GenitoriSeparati").Value
		'		Dim AffidamentoCongiunto As String = "" & Rec("AffidamentoCongiunto").Value
		'		Dim idTutore As String = "" & Rec("idTutore").Value
		'		Dim Genitore1 As String = "" & Rec("Genitore1").Value
		'		Dim Genitore2 As String = "" & Rec("Genitore2").Value
		'		Dim Genitore3 As String = "" & Rec("Genitore3").Value
		'		Dim DataDiNascita1 As String = "" & Rec("DataDiNascita1").Value
		'		Dim DataDiNascita2 As String = "" & Rec("DataDiNascita2").Value
		'		Dim DataDiNascita3 As String = "" & Rec("DataDiNascita3").Value
		'		Dim CittaDiNascita1 As String = "" & Rec("CittaNascita1").Value
		'		Dim CittaDiNascita2 As String = "" & Rec("CittaNascita2").Value
		'		Dim CittaDiNascita3 As String = "" & Rec("CittaNascita3").Value
		'		Dim Citta1 As String = "" & Rec("Citta1").Value
		'		Dim Citta2 As String = "" & Rec("Citta2").Value
		'		Dim Citta3 As String = "" & Rec("Citta3").Value
		'		Dim Indirizzo1 As String = "" & Rec("Indirizzo1").Value
		'		Dim Indirizzo2 As String = "" & Rec("Indirizzo2").Value
		'		Dim Indirizzo3 As String = "" & Rec("Indirizzo3").Value
		'		Dim Telefono1 As String = "" & Rec("TelefonoGenitore1").Value
		'		Dim Telefono2 As String = "" & Rec("TelefonoGenitore2").Value
		'		Dim Telefono3 As String = "" & Rec("Telefono3").Value
		'		Dim Mail1 As String = "" & Rec("MailGenitore1").Value
		'		Dim Mail2 As String = "" & Rec("MailGenitore2").Value
		'		Dim Mail3 As String = "" & Rec("EMail3").Value
		'		Dim Giocatore As String = "" & Rec("Giocatore").Value
		'		Rec.Close()

		'		Dim Nominativo As String = ""
		'		Dim CittaNascita As String = ""
		'		Dim DataNascita As String = ""
		'		Dim Citta As String = ""
		'		Dim Indirizzo As String = ""
		'		Dim Telefono As String = ""
		'		Dim Mail As String = ""
		'		Dim DataFirma As String = ""
		'		Dim Firma As String = ""

		'		Dim gf As New GestioneFilesDirectory
		'		Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
		'		Dim pp As String = paths
		'		pp = pp.Replace(vbCrLf, "")
		'		If Strings.Right(pp, 1) <> "\" Then
		'			pp = pp & "\"
		'		End If

		'		paths = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		'		Dim p() As String = paths.Split(";")
		'		If Strings.Right(p(0), 1) <> "\" Then
		'			p(0) = p(0) & "\"
		'		End If
		'		p(0) = p(0).Replace(vbCrLf, "")
		'		If Strings.Right(p(2), 1) <> "/" Then
		'			p(2) = p(2) & "/"
		'		End If
		'		p(2) = p(2).Replace(vbCrLf, "")

		'		Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
		'		Dim c As New CriptaFiles
		'		Dim idGenitore As Integer = 0

		'		If Maggiorenne = "S" Then
		'			Nominativo = Genitore3
		'			DataNascita = DataDiNascita3
		'			CittaNascita = CittaDiNascita3
		'			Citta = Citta3
		'			Indirizzo = Indirizzo3
		'			Telefono = Telefono3
		'			Mail = Mail3
		'			idGenitore = 3
		'		Else
		'			If GenitoriSeparati = "N" Then
		'				If Genitore1 <> "" Then
		'					Nominativo = Genitore1
		'					DataNascita = DataDiNascita1
		'					CittaNascita = CittaDiNascita1
		'					Citta = Citta1
		'					Indirizzo = Indirizzo1
		'					Telefono = Telefono1
		'					Mail = Mail1
		'					idGenitore = 1
		'				Else
		'					If Genitore2 <> "" Then
		'						Nominativo = Genitore2
		'						DataNascita = DataDiNascita2
		'						CittaNascita = CittaDiNascita2
		'						Citta = Citta2
		'						Indirizzo = Indirizzo2
		'						Telefono = Telefono2
		'						Mail = Mail2
		'						idGenitore = 2
		'					End If
		'				End If
		'			Else
		'				If AffidamentoCongiunto = "S" Then
		'					If Genitore1 <> "" Then
		'						Nominativo = Genitore1
		'						DataNascita = DataDiNascita1
		'						CittaNascita = CittaDiNascita1
		'						Citta = Citta1
		'						Indirizzo = Indirizzo1
		'						Telefono = Telefono1
		'						Mail = Mail1
		'						idGenitore = 1
		'					Else
		'						If Genitore2 <> "" Then
		'							Nominativo = Genitore2
		'							DataNascita = DataDiNascita2
		'							CittaNascita = CittaDiNascita2
		'							Citta = Citta2
		'							Indirizzo = Indirizzo2
		'							Telefono = Telefono2
		'							Mail = Mail2
		'							idGenitore = 2
		'						End If
		'					End If
		'				Else
		'					If idTutore = 1 Then
		'						Nominativo = Genitore1
		'						DataNascita = DataDiNascita1
		'						CittaNascita = CittaDiNascita1
		'						Citta = Citta1
		'						Indirizzo = Indirizzo1
		'						Telefono = Telefono1
		'						Mail = Mail1
		'						idGenitore = 1
		'					Else
		'						Nominativo = Genitore2
		'						DataNascita = DataDiNascita2
		'						CittaNascita = CittaDiNascita2
		'						Citta = Citta2
		'						Indirizzo = Indirizzo2
		'						Telefono = Telefono2
		'						Mail = Mail2
		'						idGenitore = 2
		'					End If
		'				End If
		'			End If
		'		End If

		'		Dim ddn As String = DataNascita
		'		If ddn <> "" Then
		'			Dim d() As String = ddn.Split("-")
		'			ddn = d(2) & "/" & d(1) & "/" & d(0)
		'		End If
		'		DataNascita = ddn

		'		Dim pathFirma1 As String = p(2) & NomeSquadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_4.kgb"
		'		Dim urlFirma1 As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_4.kgb"
		'		Dim pathFirmaConv1 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_4_" & Esten & ".png"
		'		Dim urlFirmaConv1 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_4_" & Esten & ".png"

		'		If ControllaEsistenzaFile(urlFirma1) Then
		'			c.DecryptFile(CryptPasswordString, urlFirma1, urlFirmaConv1)
		'			Firma = "FIRMA: <img src=""" & pathFirmaConv1 & """ style=""width: 400px; height: 150px;"" />"
		'		Else
		'			Firma = ""
		'		End If
		'		'Firma = urlFirma1

		'		Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & idgenitore
		'		Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
		'		If Not Rec.Eof() Then
		'			DataFirma = "" & Rec("DataFirma").Value
		'		Else
		'			DataFirma = ""
		'		End If
		'		Rec.Close()

		'		Dim fileDaCopiare As String = p(0) & "\" & Squadra & "\Firme\associato_" & Anno & "_" & idGiocatore & ".html"
		'		Dim fileDaCopiarePDF As String = P(0) & "\" & Squadra & "\Firme\associato_" & Anno & "_" & idGiocatore & ".pdf"
		'		Dim fileLog As String = P(0) & "\" & Squadra & "\Firme\associato_" & Anno & "_" & idGiocatore & ".log"
		'		'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
		'		gf.CreaDirectoryDaPercorso(fileDaCopiare)
		'		' Dim fileScheletro As String = Server.MapPath(".") & "\Scheletri\base_iscrizione_.txt"

		'		Dim fileScheletro As String = P(0) & Squadra & "\Scheletri\associato.txt"
		'		If Not ControllaEsistenzaFile(fileScheletro) Then
		'			fileScheletro = HttpContext.Current.Server.MapPath(".") & "\Scheletri\associato.txt"
		'		End If

		'		If ControllaEsistenzaFile(fileScheletro) Then
		'			Try
		'				Dim fileFirme As String = gf.LeggeFileIntero(fileScheletro)

		'				fileFirme = fileFirme.Replace("***nome societa menu settaggi***", NomeSquadra)
		'				fileFirme = fileFirme.Replace("***Nominativo Padre***", Nominativo)
		'				fileFirme = fileFirme.Replace("***Citta di nascita1***", CittaNascita)
		'				fileFirme = fileFirme.Replace("***Data di nascita menu anagrafica1***", DataNascita)
		'				fileFirme = fileFirme.Replace("***citta1***", Citta)
		'				fileFirme = fileFirme.Replace("****indirizzo menu anagrafica1***", Indirizzo)
		'				fileFirme = fileFirme.Replace("***telefono menu anagrafica1***", Telefono)
		'				fileFirme = fileFirme.Replace("*** mail menu anagrafica1***", Mail)
		'				fileFirme = fileFirme.Replace("***data firma2***", DataFirma)
		'				fileFirme = fileFirme.Replace("***firma padre***", Firma)
		'				fileFirme = fileFirme.Replace("***Anno menu settaggi/Dati Generali***", AnnoTabella)
		'				fileFirme = fileFirme.Replace("***Nome menu anagrafica3***", Giocatore)

		'				gf.EliminaFileFisico(fileDaCopiare)
		'				gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
		'				gf.ScriveTestoSuFileAperto(fileFirme)
		'				gf.ChiudeFileDiTestoDopoScrittura()

		'				'File.Copy(fileDaCopiare, fileDaCopiare2)
		'				Dim pp2 As New pdfGest
		'				Ritorno = pp2.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

		'				If Ritorno = "*" Then
		'					gf.EliminaFileFisico(fileDaCopiare)
		'				End If

		'				gf = Nothing
		'			Catch ex As Exception
		'				Ritorno = StringaErrore & " " & ex.Message
		'			End Try
		'		Else
		'			Ritorno = StringaErrore & " Scheletro associato non trovato"
		'		End If
		'		gf = Nothing
		'	End If
		'End If
		'End If
		'	End If
		'End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaUtenzaGenitore(Squadra As String, idGiocatore As String, Utenza As String, Cognome As String, Nome As String, Telefono As String, Mittente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim c2() As String = Squadra.Split("_")
		Dim Anno As String = Str(Val(c2(0))).Trim
		Dim codSquadra As String = c2(1)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				If TipoDB = "SQLSERVER" Then
					Sql = "Select IsNull(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti]"
				Else
					Sql = "Select Coalesce(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti]"
				End If
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Dim idUtente As String = 1

					'If Rec(0).Value Is DBNull.Value Then
					'Else
					idUtente = Rec(0).Value
					'End If
					Rec.Close()

					Dim Ok As Boolean = True
					Dim pass As String = ""
					Dim conta As Integer = 0

					While Not pass.Contains(";")
						pass = generaPassRandom()
						conta += 1
						If conta > 20 Then
							Ritorno = StringaErrore & " Creazione password fallita"
							Exit While
						End If
					End While

					If Ok Then
						Dim nuovaPass() = pass.Split(";")

						Sql = "Insert Into [Generale].[dbo].[Utenti] Values (" &
							" " & Anno & ", " &
							" " & idUtente & ", " &
							"'" & Utenza.Replace("'", "''") & "', " &
							"'" & Cognome.Replace("'", "''") & "', " &
							"'" & Nome.Replace("'", "''") & "', " &
							"'" & nuovaPass(1).Replace("'", "''") & "', " &
							"'" & Utenza.Replace("'", "''") & "', " &
							"-1, " &
							"3, " &
							" " & codSquadra & ", " &
							"1, " &
							"'" & Telefono.Replace("'", "''") & "', " &
							"'N', " &
							"';" & idGiocatore & ";', " &
							"'N', " &
							"'1-1-1-1-1' " &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno = "OK" Then
							Dim m As New mail
							Dim Oggetto As String = "Nuovo utente inCalcio"
							Dim Body As String = ""
							Body &= "E' stato creato l'utente '" & Utenza.ToUpper & "'. <br />"
							Body &= "Per accedere al sito sarà possibile digitare la mail rilasciata alla segreteria in fase di iscrizione: " & Utenza.ToUpper & "<br />"
							Body &= "La password valida per il solo primo accesso è: " & nuovaPass(0) & "<br /><br />"
							Dim ChiScrive As String = "notifiche@incalcio.cloud"

							Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, Body, Utenza, {""})
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiungeUtenzaGenitore(Squadra As String, idGiocatore As String, Utenza As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim c2() As String = Squadra.Split("_")
		Dim Anno As String = Str(Val(c2(0))).Trim
		Dim codSquadra As String = c2(1)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Update [Generale].[dbo].[Utenti] Set idGiocatore = idGiocatore + '" & idGiocatore & "'; Where Utenza = '" & Utenza.Replace("'", "''") & "'"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Ritorno = "OK" Then
					Ritorno = "*"
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class