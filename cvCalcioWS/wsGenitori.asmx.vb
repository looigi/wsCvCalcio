Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://genitori.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsGenitori
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaDatiGiocatore(Squadra As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As String = ""
				Dim Ok As Boolean = True

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value
						Rec.Close()

						Dim idGioc() As String = {}

						If idGiocatore.Contains(";") Then
							idGioc = idGiocatore.Split(";")
						Else
							idGioc = {idGiocatore}
						End If

						For Each id As String In idGioc
							If id <> "" And Ok Then
								Sql = "Select A.*, B.Descrizione From Giocatori A " &
									"Left Join [Generale].[dbo].[Ruoli] B On A.idRuolo = B.idRuolo " &
									"Where A.idGiocatore=" & id
								Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
									Ok = False
								Else
									If Rec.Eof() Then
										Ritorno = StringaErrore & " Nessun giocatore rilevato"
										Ok = False
									Else
										Ritorno &= id & ";" &
											Rec("Cognome").Value & ";" &
											Rec("Nome").Value & ";" &
											Rec("Soprannome").Value & ";" &
											Rec("DataDiNascita").Value & ";" &
											Rec("CodFiscale").Value & ";" &
											Rec("Descrizione").Value & ";" &
											"§"
										Rec.Close()
									End If
								End If
							End If
						Next
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaIdGiocatore(Squadra As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As String = ""
				Dim Ok As Boolean = True

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value
						Dim idGioc() As String = {}

						If idGiocatore.Contains(";") Then
							idGioc = idGiocatore.Split(";")
						Else
							idGioc = {idGiocatore}
						End If

						Dim gf As New GestioneFilesDirectory
						Dim path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
						Dim pp() As String = path.Split(";")
						If Strings.Right(pp(0), 1) <> "\" Then
							pp(0) = pp(0) & "\"
						End If
						Dim a() As String = Squadra.Split("_")
						Dim Anno As Integer = Val(a(0))

						For Each id As String In idGioc
							Dim p As String = pp(0) & Squadra & "\Certificati\Anno" & Anno & "\" & id & "\"
							gf.ScansionaDirectorySingola(p)
							Dim filetti() As String = gf.RitornaFilesRilevati
							Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

							Ritorno &= id & ";"
							For i As Integer = 1 To qFiletti
								Ritorno &= gf.TornaNomeFileDaPath(filetti(i)) & ";"
							Next
							Ritorno &= "§"
						Next
					End If
					Rec.Close()
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaRicevuteGiocatore(Squadra As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As String = ""
				Dim Ok As Boolean = True

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value
						Dim idGioc() As String = {}

						If idGiocatore.Contains(";") Then
							idGioc = idGiocatore.Split(";")
						Else
							idGioc = {idGiocatore}
						End If
						Rec.Close()

						Dim gf As New GestioneFilesDirectory
						Dim path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
						Dim pp() As String = path.Split(";")
						If Strings.Right(pp(0), 1) <> "\" Then
							pp(0) = pp(0) & "\"
						End If
						Dim a() As String = Squadra.Split("_")
						Dim Anno As Integer = Val(a(0))

						Ritorno = ""
						For Each id As String In idGioc
							If id <> "" Then
								Dim p As String = pp(0) & Squadra & "\Ricevute\Anno" & Anno & "\" & id & "\"
								gf.ScansionaDirectorySingola(p)
								Dim filetti() As String = gf.RitornaFilesRilevati
								Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

								For i As Integer = 1 To qFiletti
									If filetti(i).ToUpper.Contains(".PDF") Then
										Ritorno &= id & ";"
										Ritorno &= gf.TornaNomeFileDaPath(filetti(i)) & ";"
										Ritorno &= "§"
									End If
								Next
							End If
						Next

						If Ritorno = "" Then
							Ritorno = "ERROR: Nessuna ricevuta rilevata"
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMails(Squadra As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As String = ""
				Dim Ok As Boolean = True

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value
					End If
					Rec.Close()
				End If

				If Ok Then
					Dim idGioc() As String = {}

					If idGiocatore.Contains(";") Then
						idGioc = idGiocatore.Split(";")
					Else
						idGioc = {idGiocatore}
					End If

					For Each id As String In idGioc
						If id <> "" And Ok Then
							Sql = "Select A.*, B.Cognome + ' ' + B.Nome As Giocatore From GiocatoriDettaglio A " &
								"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
								"Where A.idGiocatore=" & id
							Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
							Else
								If Rec.Eof() Then
									'Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
									'Ok = False
								Else
									Dim Genitore1 As String = "" & Rec("Genitore1").Value
									Dim Genitore2 As String = "" & Rec("Genitore2").Value
									Dim Giocatore As String = "" & Rec("Giocatore").Value
									Dim mail1 As String = "" & Rec("MailGenitore1").Value
									Dim mail2 As String = "" & Rec("MailGenitore2").Value
									Dim mail3 As String = "" & Rec("MailGenitore3").Value
									Rec.Close()

									Dim Attiva1 As String = ""
									Dim Attiva2 As String = ""
									Dim Attiva3 As String = ""

									For i As Integer = 1 To 3
										Sql = "Select * From GiocatoriMails Where idGiocatore=" & id & " And Progressivo=" & i
										Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
										If TypeOf (Rec) Is String Then
											Ritorno = Rec
											Ok = False
										Else
											If Rec.Eof() Then
												Sql = "Insert Into GiocatoriMails Values (" &
													" " & id & ", " &
													" " & i & ", " &
													"'', " &
													"'N' " &
													")"
												Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
												End If
											End If
											Rec.Close()
										End If
									Next

									Sql = "Select * From GiocatoriMails Where idGiocatore=" & id
									Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
									Else
										If Not Rec.Eof() Then
											Do Until Rec.Eof()
												Select Case Val(Rec("Progressivo").Value)
													Case 1
														Attiva1 = "" & Rec("Attiva").Value
													Case 2
														Attiva2 = "" & Rec("Attiva").Value
													Case 3
														Attiva3 = "" & Rec("Attiva").Value
												End Select

												Rec.MoveNext()
											Loop
											Rec.Close()

											Ritorno &= Genitore1 & ";" & mail1 & ";" & Attiva1 & ";"
											Ritorno &= Genitore2 & ";" & mail2 & ";" & Attiva2 & ";"
											Ritorno &= Giocatore & ";" & mail3 & ";" & Attiva3 & ";"
											Ritorno &= "§"
										End If
									End If
								End If
							End If
						End If
					Next
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaMails(Squadra As String, idUtente As String, Attiva1 As String, Attiva2 As String, Attiva3 As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec as object
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As String = ""
				Dim Ok As Boolean = True

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value
					End If
					Rec.Close()
				End If

				If Ok Then
					If Not Ritorno.Contains(StringaErrore) Then
						Dim idGioc() As String = {}

						If idGiocatore.Contains(";") Then
							idGioc = idGiocatore.Split(";")
						Else
							idGioc = {idGiocatore}
						End If

						Dim sAttiva1() As String = Attiva1.Split(";")
						Dim sAttiva2() As String = Attiva2.Split(";")
						Dim sAttiva3() As String = Attiva3.Split(";")
						Dim q As Integer = 0

						For Each id As String In idGioc
							If Ok Then
								Sql = "Update GiocatoriMails Set Attiva = '" & sAttiva1(q) & "' Where idGiocatore=" & id & " And Progressivo = 1"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							End If

							If Ok Then
								Sql = "Update GiocatoriMails Set Attiva = '" & sAttiva2(q) & "' Where idGiocatore=" & id & " And Progressivo = 2"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							End If

							If Ok Then
								Sql = "Update GiocatoriMails Set Attiva = '" & sAttiva3(q) & "' Where idGiocatore=" & id & " And Progressivo = 3"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							End If

							q += 1
						Next

						Ritorno = "*"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function
End Class