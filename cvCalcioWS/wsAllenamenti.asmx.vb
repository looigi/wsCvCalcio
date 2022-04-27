Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_allti.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsAllenamenti
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function EliminaAllenamenti(Squadra As String, idAnno As String, idCategoria As String, Data As String, Ora As String, OraFine As String) As String
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
					Sql = "Delete From Allenamenti Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Datella='" & Data & "' And Orella='" & Ora & "' And OrellaFine='" & OraFine & "'"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
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
	Public Function InseriscePresenzaAllenamenti(CodiceTessera As String, NomeLettore As String, DataOra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim Conn As Object = New clsGestioneDB("Generale")

		If TypeOf (Conn) Is String Then
			Ritorno = "01" ' ErroreConnessioneDBNonValida & ":" & Conn
		Else
			Dim Rec As Object
			Dim Sql As String = "Select * From GiocatoriTessereNFC Where CodiceTessera = ' " & CodiceTessera & "'"
			Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = "02" ' Rec
			Else
				If Rec.Eof() Then
					Ritorno = "13" ' StringaErrore & " Nessuna tessera rilevata"
				Else
					Dim idGiocatore As String = Rec("idGiocatore").Value
					Dim CodSquadra As String = Rec("CodSquadra").Value
					Rec.Close()

					Ritorno = InseriscePresenzaAllenamentiFinale(CodSquadra, idGiocatore, NomeLettore, DataOra)
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function InseriscePresenzaAllenamentiFinale(Squadra As String, idGiocatore As String, NomeLettore As String, DataOra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		Dim Anno As String = ""
		Dim Mese As String = ""
		Dim Giorno As String = ""
		Dim Ora As String = ""
		Dim Minuti As String = ""
		Dim Secondi As String = ""

		If DataOra = "" Then
			Ritorno = "04" ' "ERROR: Data ora non presente"
		Else
			If DataOra.Length <> 14 Then
				Ritorno = "05" ' "ERROR: Data ora non valida"
			Else
				Dim dataValida As Boolean = True

				Try
					Anno = DataOra.Substring(0, 4)
					Mese = DataOra.Substring(4, 2)
					Giorno = DataOra.Substring(6, 2)
					Ora = DataOra.Substring(8, 2)
					Minuti = DataOra.Substring(10, 2)
					Secondi = DataOra.Substring(2, 2)

					If Val(Anno) < 2020 Or Val(Anno) > 2999 Then
						dataValida = False
					Else
						If Val(Mese) < 1 Or Val(Mese) > 12 Then
							dataValida = False
						Else
							Dim maxGiorni As Integer = 31
							If Val(Mese) = 2 Then
								If Val(Anno) / 4 = Int(Val(Anno) / 4) Then
									maxGiorni = 29
								Else
									maxGiorni = 28
								End If
							Else
								If maxGiorni = 4 Or Mese = 6 Or Mese = 9 Or Mese = 11 Then
									maxGiorni = 30
								End If
							End If
							If Val(Giorno) < 1 Or Val(Giorno) > maxGiorni Then
								dataValida = False
							Else
								If Val(Ora) < 0 Or Val(Ora) > 23 Then
									dataValida = False
								Else
									If Val(Minuti) < 0 Or Val(Minuti) > 59 Then
										dataValida = False
									Else
										If Val(Secondi) < 0 Or Val(Secondi) > 59 Then
											dataValida = False
										End If
									End If
								End If
							End If
						End If
					End If
				Catch ex As Exception
					dataValida = False
				End Try

				If dataValida = False Then
					Ritorno = "05" ' "ERROR: Data ora non valida. Inserirla nel formato yyyyMMddHHmmss"
				Else
					Dim iString As String = Anno & "-" & Mese & "-" & Giorno & " " & Ora & ":" & Minuti & ":" & Secondi
					Dim DataAttuale As DateTime = Nothing
					Dim conv As Boolean = True

					Try
						DataAttuale = DateTime.ParseExact(iString, "yyyy-MM-dd HH:mm:ss", Nothing)
					Catch ex As Exception
						conv = False
					End Try

					If conv = False Then
						Ritorno = "06" ' "ERROR: Conversione data non riuscita: " & iString
					Else
						If Connessione = "" Then
							Ritorno = ErroreConnessioneNonValida
						Else
							Dim Conn As Object = New clsGestioneDB(Squadra)

							If TypeOf (Conn) Is String Then
								Ritorno = "01" ' ErroreConnessioneDBNonValida & ":" & Conn
							Else
								Dim Rec As Object
								Dim Sql As String = ""
								Dim Ok As Boolean = True

								Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

								If Not Ritorno.Contains(StringaErrore) Then
									Sql = "Select idAnno, Categorie From Giocatori Where idGiocatore = " & idGiocatore
									Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = "02" ' Rec
									Else
										If Rec.Eof() Then
											Ritorno = "07" ' StringaErrore & " Nessun giocatore rilevato"
										Else
											Dim idAnno As String = Rec("idAnno").Value
											Dim Categorie As String = Rec("Categorie").Value
											Dim sCat() As String = {}
											If Categorie <> "" And Categorie.Contains("-") Then
												sCat = Categorie.Split("-")
											End If
											Rec.Close()

											Sql = "Select * From LettoriNFC Where Descrizione = '" & NomeLettore & "' And Eliminato='N'"
											Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
											If TypeOf (Rec) Is String Then
												Ritorno = Rec
											Else
												If Rec.Eof() Then
													Ritorno = "03" ' StringaErrore & " Nessun lettore NFC rilevato"
												Else
													Dim idLettore As String = Rec("idLettore").Value
													Rec.Close()

													For Each idCategoria As String In sCat
														Dim Progressivo As Integer = 0
														Dim GiornoAttuale As String = DataAttuale.DayOfWeek

														Sql = "Select * From Categorie Where idCategoria = " & idCategoria & " " &
															"And (GiornoAllenamento1=" & GiornoAttuale & " Or GiornoAllenamento2=" & GiornoAttuale & " Or GiornoAllenamento3=" & GiornoAttuale & ")"
														Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
														If TypeOf (Rec) Is String Then
															Ritorno = Rec
														Else
															If Rec.Eof() Then
																Ritorno = "08" ' StringaErrore & " Nessun allenamento rilevato per la categoria e il giorno attuale"
															Else
																' Dim DataAllenamentiGiorno As String = Format(Now.Date, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year
																Dim DataAllenamentiGiorno As String = Giorno & "/" & Mese & "/" & Anno
																Dim OraAllenamentiGiorno As String = ""
																Dim OraAllenamentiFine As String = ""

																If GiornoAttuale = Rec("GiornoAllenamento1").Value Then
																	OraAllenamentiGiorno = Rec("OraInizio1").Value
																	OraAllenamentiFine = Rec("OraFine1").Value
																Else
																	If GiornoAttuale = Rec("GiornoAllenamento2").Value Then
																		OraAllenamentiGiorno = Rec("OraInizio2").Value
																		OraAllenamentiFine = Rec("OraFine2").Value
																	Else
																		If GiornoAttuale = Rec("GiornoAllenamento3").Value Then
																			OraAllenamentiGiorno = Rec("OraInizio3").Value
																			OraAllenamentiFine = Rec("OraFine3").Value
																		End If
																	End If
																End If
																Rec.Close()

																If OraAllenamentiGiorno = "" Or OraAllenamentiFine = "" Then
																	Ritorno = "09" ' StringaErrore & " Nessun orario rilevato per gli allenamenti della categoria e il giorno attuale"
																Else
																	If Not OraAllenamentiGiorno.Contains(":") Or Not OraAllenamentiFine.Contains(":") Then
																		Ritorno = "10" ' StringaErrore & " Orario non valido rilevato per gli allenamenti della categoria e il giorno attuale"
																	Else
																		Dim sInizio() As String = OraAllenamentiGiorno.Split(":")
																		Dim sFine() As String = OraAllenamentiFine.Split(":")

																		Dim minutiDiAnticipoRitardo As Integer = 30
																		Dim RangeInizioOra As Integer = (Val(sInizio(0)) * 60 + Val(sInizio(1))) - minutiDiAnticipoRitardo
																		Dim RangeFineOra As Integer = (Val(sFine(0)) * 60 + Val(sFine(1))) + minutiDiAnticipoRitardo

																		Dim RangeAttualeOra As Integer = ((DataAttuale.Hour) * 60 + DataAttuale.Minute)

																		If RangeAttualeOra >= RangeInizioOra And RangeAttualeOra <= RangeFineOra Then
																			If TipoDB = " ThenSQLSERVER" Then
																				Sql = "Select IsNull(Max(Progressivo),0)+1 From Allenamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " " &
																					"And Datella=" & DataAllenamentiGiorno & " And Orella=" & OraAllenamentiGiorno
																			Else
																				Sql = "Select Coalesce(Max(Progressivo),0)+1 From Allenamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " " &
																					"And Datella=" & DataAllenamentiGiorno & " And Orella=" & OraAllenamentiGiorno
																			End If

																			Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
																			'If Rec(0).Value Is DBNull.Value Then
																			'		Progressivo = 1
																			'	Else
																			Progressivo = Rec(0).Value
																			'	End If

																			Sql = "Insert Into Allenamenti Values (" &
																				" " & idAnno & ", " &
																					" " & idCategoria & ", " &
																					"'" & DataAllenamentiGiorno & "', " &
																					"'" & OraAllenamentiGiorno & "', " &
																					" " & Progressivo & ", " &
																					" " & idGiocatore & ", " &
																					"'" & OraAllenamentiFine & "', " &
																					" " & idLettore & " " &
																					")"
																			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
																				If Ritorno.Contains(StringaErrore) Then
																					Ritorno = "12" ' "Errore nella insert"
																					Sql = "rollback"
																					Dim Ritorno3 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
																				End If
																			Else
																				Ritorno = "11" ' StringaErrore & " Orario attuale non in fascia con la categoria del giocatore"
																		End If
																	End If
																End If
															End If
														End If
													Next
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
		End If

		If Ritorno = "*" Then
			Ritorno = "00"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaAllenamenti(Squadra As String, idAnno As String, idCategoria As String, Data As String, Ora As String, Giocatori As String, OraFine As String) As String
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
					Sql = "Delete From Allenamenti Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Datella='" & Data & "' And Orella='" & Ora & "'"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						Dim sGiocatori() As String = Giocatori.Split(";")
						Dim Progressivo As Integer = 0

						For Each s As String In sGiocatori
							If s <> "" Then
								Sql = "Insert Into Allenamenti Values (" &
									" " & idAnno & ", " &
									" " & idCategoria & ", " &
									"'" & Data & "', " &
									"'" & Ora & "', " &
									" " & Progressivo & ", " &
									" " & s & ", " &
									"'" & OraFine & "', " &
									"-1 " &
									")"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Sql = "rollback"
									Dim Ritorno3 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									Exit For
								End If

								Progressivo += 1
							End If
						Next
						If Not Ritorno.Contains(StringaErrore) Then
							Ritorno = "*"

							Sql = "commit"
							Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						End If
					Else
						Sql = "rollback"
						Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					End If
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
	Public Function RitornaOreAllenamenti(Squadra As String, ByVal idAnno As String, idCategoria As String, Data As String) As String
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
					Dim sq() As String = Squadra.Split("_")
					Dim codSquadra As Integer = Val(sq(1))
					Dim NomeSquadra As String = ""
					Dim Ok As Boolean = True

					Sql = "Select Orella From Allenamenti Where Datella = '" & Data & "' And idCategoria = " & idCategoria & " Group By Orella"
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = "" ' StringaErrore & " Nessun allenamento rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("Orella").Value & ";"

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
	Public Function RitornaAllenamentiCategoriaGiorno(Squadra As String, ByVal idAnno As String, idCategoria As String, Data As String, Ora As String, Stampa As String) As String
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
					'Dim sq() As String = Squadra.Split("_")
					'Dim codSquadra As Integer = Val(sq(1))
					Dim NomeSquadra As String = ""
					Dim Ok As Boolean = True

					'If Stampa = "S" Then
					'	Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & codSquadra
					'	Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					'	If TypeOf (Rec) Is String Then
					'		Ritorno = Rec
					'		Ok = False
					'	Else
					'		If Rec.Eof() Then
					'			Ritorno = StringaErrore & " Nessuna squadra rilevata"
					'			Ok = False
					'		Else
					'			NomeSquadra = Rec("Descrizione").Value
					'		End If
					'		Rec.Close()
					'	End If
					'End If

					If Ok Then
						Dim NomeCategoria As String = ""

						If Stampa = "S" Then
							Sql = "Select * From Categorie Where idCategoria = " & idCategoria
							Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
							Else
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessuna categoria rilevata"
									Ok = False
								Else
									NomeCategoria = Rec("Descrizione").Value
								End If
								Rec.Close()
							End If
						End If

						If Ok Then
							Dim codiciGiocatore As String = ""
							Dim oraInizio As String = ""
							Dim oraFine As String = ""

							Sql = "SELECT Giocatori.idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
								"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, idCategoria2, Categorie.Descrizione As Categoria2, idCategoria3, " &
								"Cat3.Descrizione As Categoria3, Cat1.Descrizione As Categoria1, Allenamenti.Orella, Allenamenti.OrellaFine " &
								"FROM (((((Allenamenti LEFT JOIN Giocatori ON (Allenamenti.idGiocatore = Giocatori.idGiocatore) And (Allenamenti.idAnno = Giocatori.idAnno))) " &
								"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo) " &
								"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria2 And Categorie.idAnno=Giocatori.idAnno) " &
								"Left Join Categorie As Cat3 On Cat3.idCategoria=Giocatori.idCategoria3 And Cat3.idAnno=Giocatori.idAnno) " &
								"Left Join Categorie As Cat1 On Cat1.idCategoria=Giocatori.idCategoria And Cat1.idAnno=Giocatori.idAnno " &
								"Where Giocatori.Eliminato='N' And Allenamenti.idAnno=" & idAnno & " And Allenamenti.idCategoria=" & idCategoria & " And Datella='" & Data & "' And Orella='" & Ora & "' " &
								"Order By Ruoli.idRuolo, Cognome, Nome"
							Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof() Then
									Ritorno = "" ' StringaErrore & " Nessun allenamento rilevato"
								Else
									If Stampa = "N" Then
										Ritorno = ""
									Else
										Ritorno = "<table style=""width: 100%;"">"
										Ritorno &= "<tr><th>Giocatori Presenti</th></tr>"
									End If

									Dim q As Integer = 0

									Do Until Rec.Eof()
										If Stampa = "N" Then
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
											"§"
										Else
											codiciGiocatore &= Rec("idGiocatore").Value & ","
											oraInizio = Rec("Orella").Value
											oraFine = Rec("OrellaFine").Value

											q += 1
											Dim conta As String = Format(q, "00")

											Ritorno &= "<tr>"
											' Ritorno &= "<td>" & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
											Ritorno &= "<td style=""padding-left: 50px;"">" & conta & " - " & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
											Ritorno &= "</tr>"
										End If

										Rec.MoveNext()
									Loop
								End If
								Rec.Close()

								If Stampa <> "N" Then
									Ritorno &= "</table>"
									Ritorno &= "<hr />"

									If codiciGiocatore.Length > 0 Then
										codiciGiocatore = Mid(codiciGiocatore, 1, codiciGiocatore.Length - 1)
									End If

									Sql = "Select * From Giocatori " &
									"Where " & IIf(TipoDB = "SQLSERVER", "CharIndex('" & idCategoria & "-', Categorie) > 0 ", "Instr(Categorie, '" & idCategoria & "-') > 0 ") & " " &
									"And idGiocatore Not In (" & codiciGiocatore & ")"
									Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										Dim q As Integer = 0

										Ritorno &= "<table style=""width: 100%;"">"
										Ritorno &= "<tr><th>Giocatori Assenti</th></tr>"
										Do Until Rec.Eof()
											q += 1
											Dim conta As String = Format(q, "00")

											Ritorno &= "<tr>"
											Ritorno &= "<td style=""padding-left: 50px;"">" & conta & " - " & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
											Ritorno &= "</tr>"

											Rec.MoveNext()
										Loop
										Rec.Close()
										Ritorno &= "</table>"
									End If

									Dim gf As New GestioneFilesDirectory
									Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_allenamenti.txt")

									filetto = filetto.Replace("***TITOLO***", "Allenamenti " & NomeCategoria & "<br />" & Data & " " & oraInizio & "-" & oraFine)
									filetto = filetto.Replace("***DATI***", Ritorno)
									filetto = filetto.Replace("***NOME SQUADRA***", "<br /><br />" & NomeSquadra)

									Dim multimediaPaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
									Dim mmPaths() As String = multimediaPaths.Split(";")
									mmPaths(2) = mmPaths(2).Replace(vbCrLf, "")
									If Strings.Right(mmPaths(2), 1) <> "/" Then
										mmPaths(2) &= "/"
									End If

									Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
									filePaths = filePaths.Replace(vbCrLf, "")
									If Strings.Right(filePaths, 1) <> "\" Then
										filePaths &= "\"
									End If
									Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
									Dim pathLogo As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Societa\1_1.kgb"
									If ControllaEsistenzaFile(pathLogo) Then
										Dim pathLogoConv As String = filePaths & "Appoggio\" & Esten & ".jpg"
										Dim c As New CriptaFiles
										c.DecryptFile(CryptPasswordString, pathLogo, pathLogoConv)

										Dim urlLogo As String = mmPaths(2) & "Appoggio/" & Esten & ".jpg"
										filetto = filetto.Replace("***LOGO SOCIETA***", urlLogo)
									Else
										filetto = filetto.Replace("***LOGO SOCIETA***", "")
									End If

									Dim nomeFileHtml As String = filePaths & "Appoggio\" & Esten & ".html"
									Dim nomeFilePDF As String = filePaths & "Appoggio\" & Esten & ".pdf"

									gf.CreaAggiornaFile(nomeFileHtml, filetto)

									Dim pp2 As New pdfGest
									Ritorno = pp2.ConverteHTMLInPDF(nomeFileHtml, nomeFilePDF, "")
									If Ritorno = "*" Then
										Ritorno = "Appoggio/" & Esten & ".pdf"
									End If
								End If
							End If
						End If
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
	Public Function RitornaAllenamentiCategoria(Squadra As String, idCategoria As String) As String
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

				Sql = "Select Datella, Substring(Datella, 7, 4) + Substring(Datella, 4, 2) + Substring(Datella, 1, 2) As Datella2, Orella, OrellaFine From Allenamenti " &
					"Where idCategoria = " & idCategoria & " " &
					"Group By Datella, Substring(Datella, 7, 4) + Substring(Datella, 4, 2) + Substring(Datella, 1, 2), Orella, OrellaFine " &
					"Order By 2 Desc, 3"
				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = "" ' StringaErrore & " Nessun allenamento rilevato"
						Ok = False
					Else
						Ritorno = ""
						Do Until Rec.Eof()
							Ritorno &= Rec("Datella").Value & ";" & Rec("Orella").Value & ";" & Rec("OrellaFine").Value & "§"

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function
End Class