Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_uteloc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsUtentiLocali
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function ProvaWS(Cosa As String) As String
		Dim Ritorno As String = "Tiè... : " & Cosa
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ErroriLogin(idUtente As String, Errore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim idEve As Integer

				If Errore = "S" Then
					If TipoDB = "SQLSERVER" Then
						Sql = "SELECT IsNull(Max(Contatore),0)+1 FROM ErroriLogin Where idUtente=" & idUtente
					Else
						Sql = "SELECT Coalesce(Max(Contatore),0)+1 FROM ErroriLogin Where idUtente=" & idUtente
					End If
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						'If Rec(0).Value Is DBNull.Value Then
						If Rec(0).Value = 1 Then
							idEve = 1
						Else
							idEve = Rec(0).Value
						End If
						Rec.Close()

						Sql = "Select * From ErroriLogin Where idUtente=" & idUtente
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Sql = "Insert Into errorilogin Values (" & idUtente & ", " & idEve & ")"
							Else
								Sql = "Update errorilogin Set Contatore = " & idEve & " Where idUtente=" & idUtente
							End If
							Rec.Close()
						End If
					End If

					If idEve > 3 Then
						' Troppi errori. Faccio scadere la login
						Sql = "SELECT * FROM Utenti Where idUtente=" & idUtente
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim Utente As String = Rec("Utente").Value
							Ritorno = RitornaMailDopoRichiesta(Server.MapPath("."), "Generale", Utente)
							Rec.Close()

							Ritorno = "ERROR: password scaduta"
						End If
					End If
				Else
					Sql = "SELECT * FROM ErroriLogin Where idUtente=" & idUtente
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Sql = "Insert Into ErroriLogin Values (" & idUtente & ", 0)"
						Else
							Sql = "Update ErroriLogin Set Contatore = 0 Where idUtente=" & idUtente
						End If
						Rec.Close()
					End If
				End If

				If Sql <> "" Then
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaLogAccessi(idUtente As String, Descrizione As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim idEve As Integer

				If TipoDB = "SQLSERVER" Then
					Sql = "SELECT IsNull(Max(Progressivo),0)+1 As Quanti FROM LogAccessi"
				Else
					Sql = "SELECT Coalesce(Max(Progressivo),0)+1 As Quanti FROM LogAccessi"
				End If
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	idEve = 1
					'Else
					idEve = Rec(0).Value
					'End If
					Rec.Close()
				End If

				Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
				Sql = "Insert Into LogAccessi Values (" & idUtente & ", " & idEve & ", '" & Datella & "', '" & Descrizione.Replace("'", "''") & "')"

				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaUtentiGenitori(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim sq() As String = Squadra.Split("_")
				Dim idSquadra As Integer = Val(sq(1))

				Try
					' Sql = "SELECT * FROM UtentiMobile Where idAnno=" & idAnno & " And idUtente=" & idUtente
					Sql = "SELECT A.*, B.Descrizione " &
						"From [Generale].[dbo].[Utenti] A " &
						"LEFT Join Categorie B On A.idCategoria = B.idCategoria And A.idAnno = B.idAnno " &
						"Where A.idTipologia=3 And A.Eliminato='N' And A.idSquadra=" & idSquadra
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, ConnessioneGenerale)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("idAnno").Value & ";" &
									Rec("idUtente").Value & ";" &
									Rec("Utente").Value & ";" &
									Rec("Cognome").Value & ";" &
									Rec("Nome").Value & ";" &
									Rec("Password").Value & ";" &
									Rec("EMail").Value & ";" &
									Rec("idCategoria").Value & ";" &
									Rec("idTipologia").Value & ";" &
									Rec("Descrizione").Value & ";" &
									Rec("Telefono").Value & ";" &
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
	Public Function ImpostaPasswordDimenticata(ByVal Utente As String, PWD As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Sql = "SELECT Utenti.idAnno, idUtente, Utente, Cognome, Nome, " &
						"Password, EMail, idCategoria, Utenti.idTipologia, Utenti.idSquadra, Descrizione As Squadra " &
						"FROM Utenti Left Join Squadre On Utenti.idSquadra = Squadre.idSquadra " &
						"Where Upper(Utente)='" & Utente.ToUpper.Replace("'", "''") & "' And Eliminato = 'N'" ' And idAnno=" & idAnno
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
					Else
						Dim idUtente As Integer = Rec("idUtente").Value
						Dim wrapper As New CryptEncrypt(CryptPasswordString)
						Dim nuovaPassCrypt As String = wrapper.EncryptData(PWD)

						Try
							Sql = "Update Utenti Set Password='" & nuovaPassCrypt.Replace("'", "''") & "', PasswordScaduta=0 " &
									"Where idUtente=" & idUtente
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
						End Try
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMailDimenticata(ByVal Utente As String) As String
		Return RitornaMailDopoRichiesta(Server.MapPath("."), "Generale", Utente)
	End Function

	<WebMethod()>
	Public Function CreaStringaCriptata(Stringa As String) As String
		Dim wrapper As New CryptEncrypt(CryptPasswordString)
		Dim Ritorno As String = wrapper.EncryptData(Stringa)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CaricaFirme(CodSquadra As String, idSquadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(CodSquadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim NumeroFirme As Integer = 2

				Dim Sql As String = "Select * From NumeroFirme Where idSquadra=" & idSquadra
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						NumeroFirme = Rec("NumeroFirme").Value
					End If
					Rec.Close

					Dim Firme As String = ""
					Dim ConnessioneSquadra As String = LeggeImpostazioniDiBase(Server.MapPath("."), CodSquadra)
					If ConnessioneSquadra = "" Then
						Ritorno = ErroreConnessioneNonValida
					Else
						Dim ConnSquadra As Object = New clsGestioneDB(CodSquadra)

						If TypeOf (Conn) Is String Then
							Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
						Else
							Sql = "Select * From TipologiaFirme Order By idFirma"
							Rec = ConnSquadra.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSquadra)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								Do Until Rec.Eof()
									Firme &= Rec("idFirma").Value & "^" & Rec("Tipologia").value & "^" & Rec("Descrizione").Value & "!"

									Rec.MoveNext
								Loop
								Rec.Close

								Ritorno = NumeroFirme & ";" & Firme
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaUtentePerLoginNuovo(Utente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim Squadra As String = ""
		Dim UtenteDaSalvare As String = ""

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

				' Pulisce Cartella temporanea
				PulisceCartellaTemporanea(Server.MapPath("."))
				' Pulisce Cartella temporanea

				Try
					Sql = "SELECT A.idAnno, A.idUtente, Utente, Cognome, Nome, " &
						"Password, EMail, idCategoria, A.idTipologia As idTipologia, A.idSquadra, Descrizione As Squadra, PasswordScaduta, Telefono, " &
						" " & IIf(TipoDB = "SQLSERVER", "ISNULL(B.Eliminata, 'N')", "COALESCE(B.Eliminata, 'N')") & " As Eliminata, B.idTipologia As idTipo2, B.idLicenza, A.idSquadra, A.AmmOriginale, C.Mail, C.PwdMail, D.AggiornaWidgets As AggiornaWidget " &
						"FROM Utenti A Left Join Squadre B On A.idSquadra = B.idSquadra " &
						"Left Join UtentiMails C On A.idUtente = C.idUtente " &
						"Left Join AggiornamentoWidgets D On A.idSquadra = D.idSquadra " &
						"Left Join NumeroFirme E On A.idSquadra = E.idSquadra " &
						"Where Upper(Utente)='" & Utente.ToUpper.Replace("'", "''") & "' And A.Eliminato = 'N' " &
						"Order By A.idTipologia"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							'If Password <> DecriptaStringa(Rec("Password").Value.ToString) Then
							'	Ritorno = StringaErrore & " Password non valida"
							'Else
							'Dim AggiornaWidgets As Boolean = False

							Dim idSquadra As String = ""
							Dim idAnno As String = ""
							Dim codSquadra As String = ""
							Dim idLicenza As String = ""

							Ritorno = ""

							Do Until Rec.Eof()
								'AggiornaWidgets = IIf(("" & Rec("AggiornaWidget").Value) = "S", True, False)
								Dim Ok As Boolean = False

								'If Not Rec("Eliminata").Value Is DBNull.Value Then
								'	If Rec("Eliminata").Value = "N" Then
								'		Ok = True
								'	End If
								'Else
								'	Ok = True
								'End If

								'If Ok = True Then
								Dim ok2 As Boolean = True

								If Rec("idTipologia").Value = 0 Or Rec("idTipologia").Value = "0" Then
									idLicenza = 1
									idSquadra = 1
									idAnno = 1
								Else
									idLicenza = "" & Rec("idLicenza").Value
									idSquadra = Rec("idSquadra").Value.ToString.Trim
									idAnno = Rec("idAnno").Value.ToString.Trim
								End If

								Dim app1 As String = idAnno
								For i As Integer = app1.Length + 1 To 4
									app1 = "0" & app1
								Next
								Dim app2 As String = idSquadra
								For i As Integer = app2.Length + 1 To 5
									app2 = "0" & app2
								Next
								codSquadra = app1 & "_" & app2

								If Rec("idSquadra").Value <> -1 Then
									Sql = "Select * From SquadraAnni Where idSquadra=" & Rec("idSquadra").Value & " And idAnno=" & Rec("idAnno").Value
									Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
										ok2 = False
									Else
										If Rec2.Eof() Then
											Ritorno = StringaErrore & " Nessun dettaglio squadra rilevato"
											ok2 = False
										Else
											If Rec2("OnLine").Value = "N" Then
												Ritorno = StringaErrore & " La squadra dell'utente è temporanemante offline. Riprovare più tardi"
												ok2 = False
											Else
												ok2 = True
											End If

											Rec2.Close()
										End If
									End If
								End If

								If ok2 Then
									If ok2 Then
										Ritorno &= Rec("idAnno").Value & ";"
										Ritorno &= Rec("idUtente").Value & ";"
										Ritorno &= Rec("Utente").Value & ";"
										Ritorno &= Rec("Cognome").Value & ";"
										Ritorno &= Rec("Nome").Value & ";"
										Ritorno &= DecriptaStringa(Rec("Password").Value) & ";"
										Ritorno &= Rec("EMail").Value & ";"
										Ritorno &= Rec("idCategoria").Value & ";"
										Ritorno &= Rec("idTipologia").Value & ";"
										Ritorno &= Rec("idSquadra").Value & ";"
										Ritorno &= Rec("Squadra").Value & ";"
										Ritorno &= Rec("PasswordScaduta").Value & ";"
										Ritorno &= Rec("Telefono").Value & ";"
										Ritorno &= Rec("idTipo2").Value & ";"
										Ritorno &= Rec("idLicenza").Value & ";"
										Ritorno &= Rec("AmmOriginale").Value & ";"
										Ritorno &= Rec("Mail").Value & ";"
										Ritorno &= Rec("PwdMail").Value & ";"
										'Ritorno &= NumeroFirme & ";"
										'Ritorno &= Firme & ";"
										Ritorno &= "§"

										Squadra = "" & Rec("Squadra").Value
										UtenteDaSalvare = Ritorno
									End If
								End If
								'End If

								Rec.MoveNext()
							Loop
							'End If
							Rec.Close()

							If Not Ritorno.Contains(StringaErrore) Then
								Sql = "Select * From Squadre Where Eliminata = 'N' Order By Descrizione"
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									Ritorno &= "|"
									If Rec.Eof() Then
										' Ritorno = StringaErrore & " Nessuna squadra rilevata"
									Else
										Do Until Rec.Eof()
											Dim Tipologia As String = ""
											Dim Licenza As String = ""

											Select Case Rec("idTipologia").Value
												Case 1
													Tipologia = "Produzione"
												Case 2
													Tipologia = "Prova"
											End Select

											Select Case Rec("idLicenza").Value
												Case 1
													Licenza = "Base"
												Case 2
													Licenza = "Standard"
												Case 3
													Licenza = "Premium"
											End Select

											Ritorno &= Rec("idSquadra").Value & ";" & Rec("Descrizione").Value & ";" & Rec("DataScadenza").Value & ";" & Tipologia & ";" & Licenza & ";" & Rec("idTipologia").Value & ";" & Rec("idLicenza").Value & ";§"

											Rec.MoveNext()
										Loop
										Rec.Close()
									End If

									' Prende lista licenze
									Sql = "Select A.idPermesso, Descrizione, NomePerCodice From Permessi_Composizione A " &
										"Left Join Permessi_Lista B On A.idPermesso = B.idPermesso " &
										"Where A.idTipologia = " & idLicenza & " And B.Eliminato = 'N'"
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										Ritorno &= "|"
										If Rec.Eof() Then
											' Ritorno = StringaErrore & " Nessuna squadra rilevata"
										Else
											Do Until Rec.Eof()
												Ritorno &= ("" & Rec("idPermesso").Value) & ";" & ("" & Rec("Descrizione").Value) & ";" & ("" & Rec("NomePerCodice").Value) & "§"

												Rec.MoveNext
											Loop
											Rec.Close()
										End If
									End If

									'If AggiornaWidgets Then
									'	Dim w As New wsWidget

									'	w.CreaConteggi(codSquadra)
									'	w.CreaFirmeDaValidare(codSquadra, "S")
									'	w.CreaIndicatori(codSquadra)
									'	w.CreaIscritti(codSquadra)
									'	w.CreaQuoteNonSaldate(codSquadra)

									'	Sql = "Update [Generale].[dbo].[AggiornamentoWidgets] Set AggiornaWidgets='N' Where idSquadra=" & idSquadra
									'	Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									'End If
								End If
							End If
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()

				'If Not Ritorno.Contains(StringaErrore) And Squadra <> "" Then
				'	Dim Connessione2 As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra.Replace(" ", "_"))

				'	If Connessione2 = "" Then
				'		Ritorno = ErroreConnessioneNonValida
				'	Else
				'		' Dim Conn2 As Object = ApreDB(Connessione2)
				'		Dim Ritorno2 As String = ""

				'		If TypeOf (Conn) Is String Then
				'			Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
				'		Else
				'			Dim Rec2 as object
				'			Dim Sql2 As String = ""
				'			Dim Campi() As String = UtenteDaSalvare.Split(";")

				'			Sql2 = "Select * From Utenti Where Upper(Utente) = '" & Utente.ToUpper.Replace("'", "''") & "' And idAnno=" & Campi(0)
				'			Rec2 = LeggeQuery(Conn2, Sql2, Connessione2)
				'			If TypeOf (Rec) Is String Then
				'				' Ritorno = Rec2
				'			Else
				'				If Rec2.Eof() Then
				'					' Aggiungo l'utente rilevato nel db generale e non in quello di lavoro
				'					Sql2 = "Insert Into Utenti Values (" &
				'						" " & Campi(0) & ", " &
				'						" " & Campi(1) & ", " &
				'						"'" & Campi(2).Replace("'", "''") & "', " &
				'						"'" & Campi(3).Replace("'", "''") & "', " &
				'						"'" & Campi(4).Replace("'", "''") & "', " &
				'						"'" & CriptaStringa(Campi(5)).Replace("'", "''") & "', " &
				'						"'" & Campi(6).Replace("'", "''") & "', " &
				'						" " & Campi(7) & ", " &
				'						" " & Campi(8) & " " &
				'						")"
				'					Ritorno2 = EsegueSql(Conn2, Sql2, Connessione2)

				'					If Not Ritorno2.Contains(StringaErrore) Then

				'					End If
				'				End If
				'				Rec2.Close()

				'			End If
				'		End If
				'	End If
				'End If
			End If
		End If

		Return Ritorno
	End Function

	Private Function PrendeMailPWD(idAnno As String, idSquadra As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Anno As String = idAnno.Trim
		For i As Integer = Anno.Length To 3
			Anno = "0" & Anno
		Next
		Dim Squadra As String = idSquadra.Trim
		For i As Integer = Squadra.Length To 4
			Squadra = "0" & Squadra
		Next
		Dim sSquadra As String = Anno & "_" & Squadra
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), sSquadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Rec2 As Object
				Dim Sql As String = "Select * From UtentiMails Where idUtente = " & idUtente

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = ";"
					Else
						Ritorno = Rec("Mail").Value & ";" & Rec("PwdMail").Value
					End If

					Rec.Close()
				End If
			End If
		End If

		Return Ritorno
	End Function


	<WebMethod()>
	Public Function RitornaUtentePerLogin(Squadra As String, ByVal idAnno As String, Utente As String, Password As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

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
					' Sql = "SELECT * FROM Utenti Where idAnno=" & idAnno & " And Utente='" & Utente.Replace("'", "''") & "'"
					Sql = "SELECT Utenti.idAnno, idUtente, Utente, Cognome, Nome, " &
						"Password, EMail, Categorie.idCategoria As idCat1, idTipologia, Categorie.Descrizione As Descr1, Telefono " &
						"FROM (Utenti " &
						"Left Join Categorie On Utenti.idCategoria=Categorie.idCategoria And Utenti.idAnno=Categorie.idAnno) " &
						"Where Utente='" & Utente.Replace("'", "''") & "' And Utenti.idAnno=" & idAnno & " And Utenti.Eliminato='N'"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							If Password <> DecriptaStringa(Rec("Password").Value.ToString) Then
								Ritorno = StringaErrore & " Password non valida"
							Else
								Ritorno = ""
								Do Until Rec.Eof()
									Ritorno &= Rec("idAnno").Value & ";" &
										Rec("idUtente").Value & ";" &
										Rec("Utente").Value & ";" &
										Rec("Cognome").Value & ";" &
										Rec("Nome").Value & ";" &
										DecriptaStringa(Rec("Password").Value) & ";" &
										Rec("EMail").Value & ";" &
										Rec("idCat1").Value & ";" &
										Rec("idTipologia").Value & ";" &
										Rec("Descr1").Value & ";" &
										Rec("Telefono").Value & ";" &
										"§"
									Rec.MoveNext()
								Loop
							End If
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
	Public Function RitornaUtenteDaID(Squadra As String, ByVal idAnno As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

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
					' Sql = "SELECT * FROM Utenti Where idAnno=" & idAnno & " And idUtente=" & idUtente
					Sql = "SELECT Utenti.*, Categorie.Descrizione " &
						"From Utenti LEFT Join Categorie On (Utenti.idCategoria = Categorie.idCategoria) And (Utenti.idAnno = Categorie.idAnno) " &
						"Where idUtente = " & idUtente & " And Utente.Eliminato='N'"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("idAnno").Value & ";" &
									Rec("idUtente").Value & ";" &
									Rec("Utente").Value & ";" &
									Rec("Cognome").Value & ";" &
									Rec("Nome").Value & ";" &
									DecriptaStringa(Rec("Password").Value) & ";" &
									Rec("EMail").Value & ";" &
									Rec("idCategoria").Value & ";" &
									Rec("idTipologia").Value & ";" &
									Rec("Descrizione").Value & ";" &
									Rec("Telefono").Value & ";" &
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
	Public Function RitornaListaUtenti(Squadra As String, idAnno As String, Selezione As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim NomeDb As String = ""
				Dim cod() As String = Squadra.Split("_")
				Dim Anno As String = Val(cod(0)).ToString
				Dim idSquadra As String = Val(cod(1)).ToString
				Dim Altro As String = ""

				If Selezione <> "" Then
					If Strings.Right(Selezione, 1) = "," Then
						Selezione = Mid(Selezione, 1, Len(Selezione) - 1)
					End If
					Altro = "And Utenti.idTipologia In (" & Selezione & ")"
				Else
					Altro = "And Utenti=-999"
				End If

				Try
					Sql = "SELECT Utenti.idAnno, Utenti.idUtente, Utenti.Utente, Utenti.Cognome, Utenti.Nome, Utenti.EMail, Categorie.Descrizione As Categoria, " &
						"Utenti.idTipologia, Utenti.Password, Categorie.idCategoria, idSquadra, Utenti.Telefono, Utenti.AmmOriginale " &
						"FROM (Utenti LEFT JOIN [" & Squadra & "].[dbo].Categorie ON Utenti.idCategoria = Categorie.idCategoria And Utenti.idAnno = Categorie.idAnno) " &
						"Where Utenti.idAnno=" & Anno & " And Utenti.idTipologia > 0 " & Altro & " And idSquadra=" & idSquadra & " And Utenti.Eliminato='N' Order By 2,1;"
					' "Where Utenti.idAnno=" & idAnno & " Order By 2,1;"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = "" ' StringaErrore & " Nessun utente rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								'Sql = " Select * From AnnoAttualeUtenti Where idUtente=" & Rec("idUtente").Value
								'Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
								'Dim AnnoUtente As Integer = Rec("idAnno").Value
								'Rec.Close()

								'Sql = " Select * From Anni Where idAnno=" & AnnoUtente.ToString
								'Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
								'Dim NomeSquadra As String = Rec("NomeSquadra").Value
								'Rec.Close()

								Ritorno &= "0;" &
									Rec("idUtente").Value & ";" &
									Rec("Utente").Value & ";" &
									Rec("Cognome").Value & ";" &
									Rec("Nome").Value & ";" &
									Rec("EMail").Value & ";" &
									Rec("idSquadra").Value & ";" &
									Rec("idTipologia").Value & ";" &
									DecriptaStringa(Rec("Password").Value) & ";" &
									Rec("idCategoria").Value & ";" &
									Rec("Categoria").Value & ";" &
									Rec("Telefono").Value & ";" &
									Rec("AmmOriginale").Value & ";" &
									"§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()

						'Ritorno &= "£"

						'Sql = "Select * From Categorie Where idAnno=" & idAnno & " And Eliminato = 'N' Order By Ordinamento"
						'Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
						'If TypeOf (Rec) Is String Then
						'    Ritorno = Rec
						'Else
						'    If Rec.Eof() Then
						'        Ritorno = StringaErrore & " Nessuna categoria rilevata"
						'    Else
						'        Do Until Rec.Eof()
						'            Ritorno &= Rec("idCategoria").Value & ";" &
						'                Rec("Descrizione").Value & ";" &
						'                "§"
						'            Rec.MoveNext()
						'        Loop
						'    End If
						'    Rec.Close()
						'End If

						'Ritorno &= "£"
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
	Public Function RitornaNuovoID(ByVal idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim idUtente As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				'Dim idUtente As String = ""

				If TipoDB = "SQLSERVER" Then
					Sql = "SELECT IsNull(Max(idUtente),0)+1 FROM Utenti Where idAnno=" & idAnno
				Else
					Sql = "SELECT Coalesce(Max(idUtente),0)+1 FROM Utenti Where idAnno=" & idAnno
				End If
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	idUtente = "1"
					'Else
					idUtente = Rec(0).Value.ToString
					'End If
				End If
				Rec.Close()
			End If
		End If

		Return idUtente
	End Function

	<WebMethod()>
	Public Function SalvaUtente(Squadra As String, ByVal idAnno As String, idUtente As String, Utente As String, Cognome As String, Nome As String, EMail As String,
								Password As String, idCategoria As String, idTipologia As String, Telefono As String, AmmOriginale As String, Mail As String, PWD As String) As String
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
				'Dim idUtente As String = ""
				Dim Ok As Boolean = True

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Sql = "SELECT * FROM [Generale].[dbo].Utenti Where Upper(Utente)='" & Utente.Trim.ToUpper & "' And Eliminato='N'"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Try
							If idUtente <> "" Then
								Dim sq() As String = Squadra.Split("_")
								Dim idSquadra As Integer = sq(1)

								'Sql = "Select idSquadra From Squadre Where Descrizione='" & Squadra.Replace("_", " ").Replace("'", "''") & "'"
								'Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
								'If TypeOf (Rec) Is String Then
								'	Ritorno = Rec
								'Else
								'	If Rec.Eof() Then
								'		Ritorno = StringaErrore & " Nessuna squadra rilevata"
								'	Else
								'		idSquadra = Rec(0).Value
								'		Rec.Close()

								Sql = "Insert Into [Generale].[dbo].[Utenti] Values (" &
										" " & idAnno & ", " &
										" " & idUtente & ", " &
										"'" & Utente.Replace("'", "''") & "', " &
										"'" & Cognome.Replace("'", "''") & "', " &
										"'" & Nome.Replace("'", "''") & "', " &
										"'" & CriptaStringa(Password).Replace("'", "''") & "', " &
										"'" & EMail.Replace("'", "''") & "', " &
										" " & idCategoria & ", " &
										" " & idTipologia & ", " &
										" " & idSquadra & ", " &
										"0, " &
										"'" & Telefono & "', " &
										"'N', " &
										"-1, " &
										"'" & AmmOriginale & "', " &
										"'" & stringaWidgets & "' " &
										")"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

								If Ok Then
									Sql = "Insert Into [Generale].[dbo].UtentiMails Values (" &
										" " & idAnno & ", " &
										" " & idUtente & ", " &
										"'" & Mail.Replace("'", "''") & "', " &
										"'" & PWD.Replace("'", "''") & "', " &
										"'' " &
										")"
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
								End If

								'If Ok = False Then
								'	Sql = "Delete From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
								'	Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

								'	Sql = "Delete From [Generale].[dbo].UtentiMails Where idUtente=" & idUtente
								'	Ritorno2 = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								'End If
								'End If
								'End If

								'If Ritorno = "*" Then
								'	Ritorno = idUtente

								'	Dim Connessione2 As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
								'	Dim Ritorno2 As String = ""

								'	If Connessione2 = "" Then
								'		Ritorno2 = ErroreConnessioneNonValida
								'	Else
								'		' Dim Conn2 As Object = ApreDB(Connessione2)

								'		If TypeOf (Conn2) Is String Then
								'			Ritorno2 = ErroreConnessioneDBNonValida & ":" & Conn2
								'		Else
								'			Dim Rec2 as object
								'			Dim Sql2 As String = ""

								'			Sql2 = "Insert Into Utenti Values (" &
								'				" " & idAnno & ", " &
								'				" " & idUtente & ", " &
								'				"'" & Utente.Replace("'", "''") & "', " &
								'				"'" & Cognome.Replace("'", "''") & "', " &
								'				"'" & Nome.Replace("'", "''") & "', " &
								'				"'" & CriptaStringa(Password).Replace("'", "''") & "', " &
								'				"'" & EMail.Replace("'", "''") & "', " &
								'				" " & idCategoria & ", " &
								'				" " & idTipologia & " " &
								'				")"
								'			Ritorno2 = EsegueSql(Conn2, Sql2, Connessione2)

								'			If Ritorno2<> "OK" Then
								'				Ritorno = Ritorno2
								'			End If
								'		End If
								'	End If
								'End If
							End If
						Catch ex As Exception
							Ok = False
							Ritorno = StringaErrore & " " & ex.Message
						End Try
					Else
						Ok = False
						Ritorno = StringaErrore & " Utente già esistente"
					End If
				End If

				If Ok = True Then
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
	Public Function EliminaUtente(Squadra As String, ByVal idAnno As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Sql = "Update Utenti Set Eliminato = 'S' Where idUtente=" & idUtente & " And idAnno=" & idAnno
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaUtente(Squadra As String, ByVal idAnno As String, Utente As String, Cognome As String, Nome As String, EMail As String,
								Password As String, idCategoria As String, idTipologia As String, idUtente As String, Telefono As String, AmmOriginale As String,
								   Mail As String, PWD As String) As String
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

				'Sql = "Delete From Utenti Where idUtente=" & idUtente & " And idAnno=" & idAnno
				'Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				'If Ritorno.Contains(StringaErrore) Then
				'	Ok = False
				'End If

				If Ok Then
					'Dim idSquadra As Integer
					'Dim sq() As String = Squadra.Split("_")

					Try
						' "idSquadra=" & Val(sq(1)).ToString & ", " &
						Sql = "Update [Generale].[dbo].[Utenti] Set " &
							"idAnno=" & idAnno & ", " &
							"Utente='" & Utente.Replace("'", "''") & "', " &
							"Cognome='" & Cognome.Replace("'", "''") & "', " &
							"Nome='" & Nome.Replace("'", "''") & "', " &
							"Password='" & CriptaStringa(Password).Replace("'", "''") & "', " &
							"EMail='" & EMail.Replace("'", "''") & "', " &
							"idCategoria=" & idCategoria & ", " &
							"idTipologia=" & idTipologia & ", " &
							"PasswordScaduta=0, " &
							"Telefono='" & Telefono & "', " &
							"Eliminato='N', " &
							"idGiocatore=-1 " &
							"Where idUtente=" & idUtente
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						Else
							Sql = "Select * From [Generale].[dbo].[UtentiMails] Where idUtente = " & idUtente
							Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
							Else
								If Rec.Eof() Then
									Sql = "Insert Into [Generale].[dbo].[UtentiMails] Values(" & idAnno & ", " & idUtente & ", '" & Mail.Replace("'", "''") & "', '" & PWD.Replace("'", "''") & "', '')"
								Else
									Sql = "Update [Generale].[dbo].[UtentiMails] Set Mail='" & Mail.Replace("'", "''") & "', PwdMail='" & PWD.Replace("'", "''") & "' Where idAnno=" & idAnno & " And idUtente=" & idUtente
								End If
								Rec.Close()

								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								Else
									Ritorno = idUtente
								End If
							End If
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
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
			'		End If
		End If

		Return Ritorno
	End Function

End Class