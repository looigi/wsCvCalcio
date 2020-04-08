Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_gioc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsGiocatori
	Inherits System.Web.Services.WebService

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
						"Order By Ruoli.idRuolo, Cognome, Nome"
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
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And CharIndex('" & idCategoria & "-', Giocatori.idCategoria) = 0 " &
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
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " " &
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
											TotalePagamento As String, TelefonoGenitore1 As String, TelefonoGenitore2 As String) As String
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
							"TelefonoGenitore2='" & Replace(TelefonoGenitore2, "'", "''") & "' " &
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
				Dim Sql As String = ""

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
							"'' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Not Ritorno.Contains(StringaErrore) Then
							Ritorno = ";;N;N;N;0;;;"
						End If
					Else
						Ritorno = Rec("Genitore1").Value & ";"
						Ritorno &= Rec("Genitore2").Value & ";"
						Ritorno &= Rec("FirmaGenitore1").Value & ";"
						Ritorno &= Rec("FirmaGenitore2").Value & ";"
						Ritorno &= Rec("CertificatoMedico").Value & ";"
						Ritorno &= Rec("TotalePagamento").Value & ";"
						Ritorno &= Rec("TelefonoGenitore1").Value & ";"
						Ritorno &= Rec("TelefonoGenitore2").Value & ";"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaGiocatore(Squadra As String, idAnno As String, idCategoria As String, idGiocatore As String, idRuolo As String, Cognome As String, Nome As String, EMail As String, Telefono As String,
								   Soprannome As String, DataDiNascita As String, Indirizzo As String, CodFiscale As String, Maschio As String, Citta As String, Matricola As String,
								   NumeroMaglia As String, idCategoria2 As String, idCategoria3 As String, Categorie As String, RapportoCompleto As String) As String
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
							"0, " &
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
	Public Function SalvaPagamento(Squadra As String, idAnno As String, idGiocatore As String, Pagamento As String) As String
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
							"'N'" &
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
							Ritorno &= Rec("Progressivo").Value & ";" & Rec("Pagamento").Value & ";" & Rec("DataPagamento").Value & ";§"
							Totale += (Rec("Pagamento").Value)

							Rec.MoveNext
						Loop
						Rec.Close
						Ritorno &= "Totale;" & Format(Totale, "#0.#0") & ";;§"
						Dim Differenza As Single = TotPag - Totale
						Differenza = CInt(Differenza * 100) / 100
						Ritorno &= "Differenza;" & Format(Differenza, "#0.#0") & ";;§"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class