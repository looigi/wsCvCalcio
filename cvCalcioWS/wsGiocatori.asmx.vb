﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices

<System.Web.Services.WebService(Namespace:="http://cvcalcio_gioc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsGiocatori
	Inherits System.Web.Services.WebService

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
	Public Function RichiedeFirma(Squadra As String, ByVal idGiocatore As String, ByVal Genitore As String) As String
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

				'Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				'Rec = LeggeQuery(Conn, Sql, Connessione)
				'If TypeOf (Rec) Is String Then
				'	Ritorno = Rec
				'Else
				'	If Not Rec.Eof Then
				'		Dim Datella As String = Rec("Datella").Value
				'		Rec.Close

				'		Dim Date1 As Date = Now
				'		Dim Date2 As Date = Date.Parse(Datella)
				'		Dim days As Long = Math.Abs(DateDiff(DateInterval.Day, Date1, Date2))
				'		If days < 5 Then
				'			If Genitore <> 3 Then
				'				Ritorno = StringaErrore & " Una firma è già stata richiesta per il giocatore ed il genitore in data " & Datella
				'			Else
				'				Ritorno = StringaErrore & " Una firma è già stata richiesta per il giocatore in data " & Datella
				'			End If
				'		Else
				Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				Ritorno = EsegueSql(Conn, Sql, Connessione)
				If Ritorno <> "*" Then
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)

					Return Ritorno
				End If
				'		End If
				'	Else
				'		Rec.Close
				'	End If

				' If Ritorno = "*" Then
				Sql = "Insert Into GiocatoriFirme Values (" & idGiocatore & ", " & Genitore & ", '" & Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00") & "', '')"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

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
											Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsoSito.txt")
											gf = Nothing

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
													Body &= "E' stata richiesta la firma di " & nomeGenitore & " dalla direzione della società " & NomeSquadra & " per l'iscrizione all'anno " & Descrizione & ". " & vbCrLf & vbCrLf
													Body &= "Per effettuare l'operazione eseguire il seguente link:" & vbCrLf & vbCrLf
													Body &= Percorso & "?firma=true&codSquadra=" & Squadra & "&id=" & idGiocatore & "&squadra=" & NomeSquadra.Replace(" ", "_") & "&anno=" & Anno & "&genitore=" & Genitore
												Else
													Body &= "E' stata richiesta la firma di " & nomeGenitore & " dalla direzione della società " & NomeSquadra & " per l'iscrizione all'anno " & Descrizione & " del giocatore " & Nominativo & ". " & vbCrLf & vbCrLf
													Body &= "Per effettuare l'operazione eseguire il seguente link:" & vbCrLf & vbCrLf
													Body &= Percorso & "?firma=true&codSquadra=" & Squadra & "&id=" & idGiocatore & "&squadra=" & NomeSquadra.Replace(" ", "_") & "&anno=" & Anno & "&genitore=" & Genitore
												End If
												Dim ChiScrive As String = "notifiche@incalcio.cloud"

												Ritorno = m.SendEmail(Oggetto, Body, ChiScrive, EMail)
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

				Try
					Sql = "SELECT idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, idCategoria2, Categorie.Descrizione As Categoria2, idCategoria3, Cat3.Descrizione As Categoria3, Cat1.Descrizione As Categoria1, " &
						"Giocatori.Categorie, Giocatori.RapportoCompleto " &
						"FROM (((Giocatori " &
						"Left Join Ruoli On Giocatori.idRuolo=Ruoli.idRuolo) " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria2 And Categorie.idAnno=Giocatori.idAnno) " &
						"Left Join Categorie As Cat3 On Cat3.idCategoria=Giocatori.idCategoria3 And Cat3.idAnno=Giocatori.idAnno) " &
						"Left Join Categorie As Cat1 On Cat1.idCategoria=Giocatori.idCategoria And Cat1.idAnno=Giocatori.idAnno " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And CharIndex('" & idCategoria & "-', Categorie) > 0 " &
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
						"Giocatori.Categorie, Giocatori.RapportoCompleto " &
						"FROM Giocatori " &
						"Left Join Ruoli On Giocatori.idRuolo=Ruoli.idRuolo " &
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
						"Giocatori.Categorie, Giocatori.RapportoCompleto " &
						"FROM Giocatori " &
						"Left Join Ruoli On Giocatori.idRuolo=Ruoli.idRuolo " &
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
						nomesquadra = Rec("NomeSquadra").Value
					End If
				End If
				Rec.Close

				If Ritorno = "" Then
					Dim gf As New GestioneFilesDirectory
					Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PathCvCalcio.txt")
					gf = Nothing
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "\" Then
						Percorso &= "\"
					End If

					Try
						Sql = "SELECT Giocatori.idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
							"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, Giocatori.idCategoria2 As idCategoria2, Categorie2.Descrizione As Categoria2, " &
							"Giocatori.idCategoria3 As idCategoria3, Categorie3.Descrizione As Categoria3, Categorie.Descrizione As Categoria1, Giocatori.Categorie, " &
							"Giocatori.RapportoCompleto, Giocatori.idTaglia, Min(KitGiocatori.idTipoKit) As idTipologiaKit " &
							"FROM Giocatori " &
							"Left Join KitGiocatori On Giocatori.idGiocatore=KitGiocatori.idGiocatore " &
							"Left Join Ruoli On Giocatori.idRuolo=Ruoli.idRuolo " &
							"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
							"Left Join Categorie As Categorie2 On Categorie2.idCategoria=Giocatori.idCategoria2 And Categorie2.idAnno=Giocatori.idAnno " &
							"Left Join Categorie As Categorie3 On Categorie3.idCategoria=Giocatori.idCategoria3 And Categorie3.idAnno=Giocatori.idAnno " &
							"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " " &
							"Group By Giocatori.idGiocatore, Ruoli.idRuolo, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, CodFiscale, Maschio, " &
							"Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, Giocatori.idCategoria2, Categorie2.Descrizione, Giocatori.idCategoria3, Categorie3.Descrizione, Categorie.Descrizione, Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.idTaglia " &
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
										Ritorno = Rec
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
									Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & Rec("idGiocatore").Value & "_1.png"
									Dim path2 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & Rec("idGiocatore").Value & "_2.png"
									Dim path3 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & Rec("idGiocatore").Value & "_3.png"
									Dim q As Integer = 0
									If File.Exists(path1) Then
										q += 1
									End If
									If File.Exists(path2) Then
										q += 1
									End If
									If File.Exists(path3) Then
										q += 1
									End If
									If q = 3 Then
										Semaforo3 = "verde"
										Titolo3 = "Firme complete"
									Else
										If q > 0 Then
											Semaforo3 = "giallo"
											Titolo3 = "Firme non complete (" & q & "/3)"
										Else
											Semaforo3 = "rosso"
											Titolo3 = "Nessuna firma presente"
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

													If Rec2("CertificatoMedico").Value = "S" And dat > Now Then
														Semaforo4 = "verde"
														Titolo4 = "Certificato e data scadenza presenti"
													Else
														Semaforo4 = "giallo"
														Titolo4 = "Certificato presente e data scaduta"
													End If
												End If
											End If
										Else
											Semaforo4 = "rosso"
											Titolo4 = "Nessun certificato e data presenti"
										End If
										Rec2.Close
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
										Rec("idTipologiaKit").Value.ToString & ";" &
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
						"Giocatori.RapportoCompleto " &
						"FROM Giocatori " &
						"Left Join Ruoli On Giocatori.idRuolo=Ruoli.idRuolo " &
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
											ScadenzaCertificatoMedico As String, MailGenitore1 As String, MailGenitore2 As String, FirmaGenitore3 As String, MailGenitore3 As String) As String
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
							"MailGenitore3='" & MailGenitore3 & "' " &
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
									"'' " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Not Ritorno.Contains(StringaErrore) Then
									Ritorno = ";;N;N;N;0;;;"
								End If
							Else
								Dim gf As New GestioneFilesDirectory
								Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PathCvCalcio.txt")
								gf = Nothing
								Percorso = Percorso.Trim()
								If Strings.Right(Percorso, 1) <> "\" Then
									Percorso &= "\"
								End If
								Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_1.png"
								Dim path2 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_2.png"
								Dim path3 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_3.png"
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
								   idTaglia As String) As String
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
					If idGiocatore = "-1" Then
						Sql = "SELECT * FROM Giocatori Where idAnno=" & idAnno & " And CodFiscale=" & CodFiscale
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Not Rec.Eof Then
							Ritorno = "Codice fiscale già presente in archivio"
						Else
							Ok = False
						End If
						Rec.Close

						If Ok Then
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
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If
					Else
						Sql = "SELECT * FROM Giocatori Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Not Rec.Eof Then
							Dim conta As Integer = 0

							'Do While Ritorno.Contains(StringaErrore) Or Ritorno = ""
							Try
								Sql = "Delete  From Giocatori Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
								'Exit Do
							End Try
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
					End If

					If Ok = True Then
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
							"'" & RapportoCompleto & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
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
	Public Function SalvaPagamento(Squadra As String, idAnno As String, idGiocatore As String, Pagamento As String, Commento As String) As String
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

					Try
						Sql = "SELECT Max(Progressivo)+1 FROM GiocatoriPagamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Rec(0).Value Is DBNull.Value Then
							Progressivo = 1
						Else
							Progressivo = Rec(0).Value
						End If
						Rec.Close()

						Dim DataPagamento As String = Now.Year & "-" & Format(Now.Month, "00") & "-" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

						Sql = "Insert Into GiocatoriPagamenti Values (" &
							" " & idAnno & ", " &
							" " & idGiocatore & ", " &
							" " & Progressivo & ", " &
							" " & Pagamento & ", " &
							"'" & DataPagamento & "', " &
							"'N', " &
							"'" & Commento.Replace("'", "''") & "' " &
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
	Public Function EliminaPagamentoGiocatore(Squadra As String, idAnno As String, idGiocatore As String, Progressivo As String) As String
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
							Ritorno = "Totale a pagare;" & Format(TotPag, "#0.#0") & ";;§"

							Do Until Rec.Eof
								Ritorno &= Rec("Progressivo").Value & ";" & Rec("Pagamento").Value & ";" & Rec("DataPagamento").Value & ";" & Rec("Commento").Value & ";§"
								Totale += (Rec("Pagamento").Value)

								Rec.MoveNext
							Loop
							Rec.Close
							Ritorno &= "Totale;" & Format(Totale, "#0.#0") & ";;§"
							Dim Differenza As Single = TotPag - Totale
							Differenza = CInt(Differenza * 100) / 100
							Ritorno &= "Differenza;" & Format(Differenza, "#0.#0") & ";;§"
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
					Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PathCvCalcio.txt")
					gf = Nothing
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "\" Then
						Percorso &= "\"
					End If
					Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_" & idGenitore & ".png"
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

			End If
		End If

		Return Ritorno
	End Function

End Class