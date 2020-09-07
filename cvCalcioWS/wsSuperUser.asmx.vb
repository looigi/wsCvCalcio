Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.IO
Imports System.Linq

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://cvcalcio_super_user.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsSuperUser
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function CreaDB(Squadra As String, DataScadenza As String, MailAdmin As String, NomeAdmin As String, CognomeAdmin As String, Anno As String, idTipologia As String,
						   idLicenza As String, Mittente As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim ConnessioneDBVuoto As String = LeggeImpostazioniDiBase(Server.MapPath("."), "DBVUOTO")
		Dim nomeDb As String = ""

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = ApreDB(ConnessioneGenerale)
			Dim ConnDbVuoto As Object = ApreDB(ConnessioneDBVuoto)
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Or TypeOf (ConnDbVuoto) Is String Then
				If TypeOf (ConnGen) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
				Else
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnDbVuoto
				End If
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim gf As New GestioneFilesDirectory

				gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Log\")
				gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "\Log\CreazioneSocieta_" & Squadra & "_" & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".txt")
				gf.ScriveTestoSuFileAperto("Squadra: " & Squadra)
				gf.ScriveTestoSuFileAperto("DataScadenza: " & DataScadenza)
				gf.ScriveTestoSuFileAperto("MailAdmin: " & MailAdmin)
				gf.ScriveTestoSuFileAperto("NomeAdmin: " & NomeAdmin)
				gf.ScriveTestoSuFileAperto("CognomeAdmin: " & CognomeAdmin)
				gf.ScriveTestoSuFileAperto("Anno: " & Anno)
				gf.ScriveTestoSuFileAperto("idTipologia: " & idTipologia)
				gf.ScriveTestoSuFileAperto("idLicenza: " & idLicenza)
				gf.ScriveTestoSuFileAperto("Mittente: " & Mittente)
				gf.ScriveTestoSuFileAperto("-------------------------------------------")

				Sql = "Select Max(idSquadra)+1 From Squadre"
				Rec = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					Dim idSocieta As Integer = -1

					If Rec(0).Value Is DBNull.Value = True Then
						idSocieta = 1
					Else
						idSocieta = Rec(0).Value
						Rec.Close
					End If
					gf.ScriveTestoSuFileAperto("idSocieta: " & idSocieta.ToString)

					Sql = "Select Max(idUtente)+1 From Utenti"
					Rec = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Dim idUtente As Integer = -1

						If Rec(0).Value Is DBNull.Value = True Then
							idUtente = 1
						Else
							idUtente = Rec(0).Value
							Rec.Close
						End If
						gf.ScriveTestoSuFileAperto("idUtente: " & idUtente.ToString)

						If Ok Then
							Dim Tabelle(0) As String
							Dim q As Integer = 0

							Sql = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'"
							Rec = LeggeQuery(ConnDbVuoto, Sql, ConnessioneDBVuoto)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
							Else
								Do Until Rec.Eof()
									ReDim Preserve Tabelle(q)
									Tabelle(q) = Rec("TABLE_NAME").Value

									q += 1
									Rec.MoveNext()
								Loop
								Rec.Close()
								gf.ScriveTestoSuFileAperto("Tabelle: " & q)
							End If

							If Ok Then
								Dim Societa As String = Format(idSocieta, "00000")
								nomeDb = "0001_" & Societa

								gf.ScriveTestoSuFileAperto("idSocieta 2: " & Societa)
								gf.ScriveTestoSuFileAperto("Nome DB: " & nomeDb)

								Sql = "Create Database [" & nomeDb & "]"
								Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
									gf.ScriveTestoSuFileAperto(Ritorno & " -> " & Sql)
								Else
									gf.ScriveTestoSuFileAperto("DB Creato")
								End If

								If Ok Then
									Sql = "Begin transaction"
									Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

									Sql = "Insert Into Squadre Values (" & idSocieta & ", '" & Squadra.Replace("'", "''") & "', '" & DataScadenza & "', " & idTipologia & ", " & idLicenza & ", 'N', " & Now.Month & ", " & Now.Year & ")"
									Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
										gf.ScriveTestoSuFileAperto(Ritorno & " ->" & Sql)
									Else
										gf.ScriveTestoSuFileAperto("Inserita riga tabella Squadre")
									End If

									If Ok Then
										Dim pass As String = generaPassRandom()
										Dim nuovaPass() = pass.Split(";")

										Sql = "Insert Into Utenti Values (" &
											"1, " &
											" " & idUtente & ", " &
											"'" & MailAdmin.Replace("'", "''") & "', " &
											"'" & CognomeAdmin.Replace("'", "''") & "', " &
											"'" & NomeAdmin.Replace("'", "''") & "', " &
											"'" & nuovaPass(1).Replace("'", "''") & "', " &
											"'" & MailAdmin.Replace("'", "''") & "', " &
											"-1, " &
											"1, " &
											" " & idSocieta & ", " &
											"1, " &
											"'', " &
											"'N', " &
											"-1, " &
											"'S', " &
											"'" & stringaWidgets & "' " &
											")"
										Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
											gf.ScriveTestoSuFileAperto(Ritorno & " -> " & Sql)
										Else
											gf.ScriveTestoSuFileAperto("Inserita riga tabella Squadre")
										End If

										If Ok Then
											gf.ScriveTestoSuFileAperto("Utente Inserito")
											For i As Integer = 0 To q - 1
												Try
													Sql = "Select * Into [" & nomeDb & "].[dbo].[" & Tabelle(i) & "] From " & Tabelle(i)
													Ritorno = EsegueSql(ConnDbVuoto, Sql, ConnessioneDBVuoto)
													If Ritorno.Contains(StringaErrore) Then
														Ok = False
														gf.ScriveTestoSuFileAperto("Errore su copia Tabella: " & Tabelle(i) & " -> " & Ritorno)
														Exit For
													Else
														gf.ScriveTestoSuFileAperto("Copiata Tabella: " & Tabelle(i))
													End If
												Catch ex As Exception
													Ritorno = StringaErrore & " " & ex.Message
													Ok = False
												End Try
											Next

											If Ok Then
												Dim ConnessioneNuovo As String = LeggeImpostazioniDiBase(Server.MapPath("."), nomeDb)
												Dim ConnNuovo As Object = ApreDB(ConnessioneNuovo)

												If TypeOf (ConnNuovo) Is String Then
													Ritorno = ErroreConnessioneDBNonValida & ":" & ConnNuovo
													Ok = False
												Else
													For i As Integer = 0 To q - 1
														Sql = "Select Chiave From _CHIAVI_ Where Upper(lTrim(rTrim(Tabella)))='" & Tabelle(i).Trim.ToUpper & "'"
														Rec = LeggeQuery(ConnNuovo, Sql, ConnessioneNuovo)
														If TypeOf (Rec) Is String Then
															Ritorno = Rec
															Ok = False
															Exit For
														Else
															If Not Rec.Eof() Then
																Dim Query As String = Rec(0).Value

																Ritorno = EsegueSql(ConnNuovo, Query, ConnessioneNuovo)
																If Ritorno.Contains(StringaErrore) Then
																	gf.ScriveTestoSuFileAperto("Errore su creazione Chiave: " & Tabelle(i) & " -> " & Ritorno)
																	Ok = False
																	Exit For
																Else
																	gf.ScriveTestoSuFileAperto("Creata Chiave: " & Tabelle(i))
																End If

																Rec.Close()
															End If
														End If
													Next

													If Ok Then
														Sql = "Drop Table _CHIAVI_"
														Ritorno = EsegueSql(ConnNuovo, Sql, ConnessioneNuovo)
														If Ritorno.Contains(StringaErrore) Then
															Ok = False
															gf.ScriveTestoSuFileAperto("Eliminazione tabella _CHIAVI_: " & Ritorno)
														Else
															gf.ScriveTestoSuFileAperto("Eliminata tabella _CHIAVI_")
														End If
													End If
												End If
											End If

											If Ok = True Then
												Sql = "Insert Into [" & nomeDb & "].[dbo].[Anni] Values (" &
													"1, " &
													"'" & Anno & "', " &
													"'" & Squadra.Replace("'", "''") & "', " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null, " &
													"null " &
													")"
												Ritorno = EsegueSql(ConnDbVuoto, Sql, ConnessioneDBVuoto)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
													gf.ScriveTestoSuFileAperto("Creazione voce su Tabella Anni: " & Ritorno)
												Else
													gf.ScriveTestoSuFileAperto("Dati inseriti in tabella Anni: " & "[" & nomeDb & "].[dbo].[Anni]")
													Sql = "Insert Into SquadraAnni Values (" &
														" " & idSocieta & ", " &
														"1, " &
														"'" & Anno & "', " &
														"'S' " &
														")"
													Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
													If Ritorno.Contains(StringaErrore) Then
														Ok = False
														gf.ScriveTestoSuFileAperto(Ritorno & " -> " & Sql)
													Else
														gf.ScriveTestoSuFileAperto("Dati inseriti in tabella SquadraAnni")
														Try
															Dim Tipologia As String = ""
															Dim Licenza As String = ""

															Select Case idTipologia
																Case 1
																	Tipologia = "Produzione"
																Case 2
																	Tipologia = "Prova"
															End Select

															Select Case idLicenza
																Case 1
																	Licenza = "Base"
																Case 2
																	Licenza = "Standard"
																Case 3
																	Licenza = "Premium"
															End Select

															Dim m As New mail
															Dim Oggetto As String = "Creazione nuova società"
															Dim Body As String = ""

															Body &= "E' stata creata la nuova società '" & Squadra & "'. <br /><br />"
															Body &= "Amministratore: " & CognomeAdmin & " " & NomeAdmin & "<br />"
															Body &= "Scadenza licenza: " & DataScadenza & "<br />"
															Body &= "Anno: " & Anno & "<br />"
															Body &= "Tipologia: " & Tipologia & "<br />"
															Body &= "Licenza: " & Licenza & "<br /><br />"
															Body &= "Per accedere, l'amministratore potrà utilizzare la password " & nuovaPass(0) & " che dovrà essere modificata al primo ingresso."

															Dim ChiScrive As String = "servizioclienti@incalcio.cloud"

															gf.ScriveTestoSuFileAperto("Invio Mail")

															Ritorno = m.SendEmail(Squadra, Mittente, Oggetto, Body, MailAdmin, {""})
															If Ritorno.Contains(StringaErrore) Then
																gf.ScriveTestoSuFileAperto("Ritorno invio mail: " & Ritorno)
															Else
																Ritorno = Societa
															End If
														Catch ex As Exception
															Ritorno = StringaErrore & " " & ex.Message
															gf.ScriveTestoSuFileAperto("Errore invio mail: " & Ritorno)
														End Try
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

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Commit: " & Ritorno2)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Rollback: " & Ritorno2)

					ConnGen.Close
					ConnDbVuoto.close

					Sql = "Drop Database [" & nomeDb & "]"
					Ritorno2 = EsegueSql(ConnDbVuoto, Sql, ConnessioneDBVuoto)
					gf.ScriveTestoSuFileAperto("Drop Database: " & Ritorno2)
				End If

				gf.ChiudeFileDiTestoDopoScrittura()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaSquadre() As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = ApreDB(ConnessioneGenerale)
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Sql = "Select * From Squadre Where Eliminata = 'N' Order By Descrizione"
				Rec = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						Do Until Rec.Eof
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

							Dim Scadenza As String = "" & Rec("DataScadenza").Value
							Dim Semaforo1 As String = "" : Dim Titolo1 As String = ""
							Dim Semaforo2 As String = "" : Dim Titolo2 As String = ""

							If Scadenza <> "" Then
								Dim sc() As String = Scadenza.Split("-")
								Scadenza = sc(2) & "/" & sc(1) & "/" & sc(0)
								Dim dScadenza As DateTime = Convert.ToDateTime(Scadenza)
								Dim Oggi As Date = Now
								Dim diff As Integer = DateAndTime.DateDiff(DateInterval.Day, Oggi, dScadenza, )

								Select Case diff
									Case < 0
										Semaforo1 = "rosso"
									Case 0 To 30
										Semaforo1 = "rosso"
									Case 31 To 90
										Semaforo1 = "giallo"
									Case > 90
										Semaforo1 = "verde"
								End Select
								Titolo1 = "Giorni alla scadenza: " & diff
							End If

							Dim Anni As Integer = 0
							Dim maxAnno As String = ""
							Dim Stato As String = ""

							Sql = "Select Count(*) From SquadraAnni Where idSquadra = " & Rec("idSquadra").Value
							Rec2 = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2
							Else
								If Not Rec2(0).Value Is DBNull.Value Then
									Anni = Rec2(0).Value
								End If
								Rec2.Close
							End If

							Sql = "Select Top 1 * From SquadraAnni Where idSquadra = " & Rec("idSquadra").Value & " Order By idAnno Desc"
							Rec2 = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2
							Else
								If Not Rec2.Eof() Then
									maxAnno = Rec2("idAnno").Value
									Stato = Rec2("OnLine").Value
								End If
								Rec2.Close
							End If

							If Stato = "S" Then
								Semaforo2 = "verde" : Titolo2 = "Database in linea"
							Else
								Semaforo2 = "rosso" : Titolo2 = "Database fuori linea"
							End If

							Ritorno &= Rec("idSquadra").Value & ";" &
									Rec("Descrizione").Value & ";" &
									Rec("DataScadenza").Value & ";" &
									Tipologia & ";" &
									Licenza & ";" &
									Semaforo1 & "*" & Titolo1 & ";" &
									Rec("idTipologia").Value & ";" &
									Rec("idLicenza").Value & ";" &
									Anni & ";" &
									maxAnno & ";" &
									Semaforo2 & "*" & Titolo2 & ";" &
									"§"

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaSquadra(idSquadra As String, Squadra As String, DataScadenza As String, idTipologia As String, idLicenza As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = ApreDB(ConnessioneGenerale)
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Try
					Sql = "Update Squadre Set " &
						"Descrizione='" & Squadra.Replace("'", "''") & "'," &
						"DataScadenza='" & DataScadenza & "'," &
						"idTipologia=" & idTipologia & "," &
						"idLicenza=" & idLicenza & " " &
						"Where idSquadra=" & idSquadra
					Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
					If Not Ritorno.Contains(StringaErrore) Then
						Ritorno = "*"
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaSquadra(idSquadra As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = ApreDB(ConnessioneGenerale)
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Try
					Sql = "Update Squadre Set " &
						"Eliminata='S'" &
						"Where idSquadra=" & idSquadra
					Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
					If Not Ritorno.Contains(StringaErrore) Then
						Ritorno = "*"
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaNuovoAnno(Squadra As String, idSquadra As String, NuovoAnno As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim ConnessioneDBOrigine As String = LeggeImpostazioniDiBase(Server.MapPath("."), "DBVUOTO")

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = ApreDB(ConnessioneGenerale)
			Dim ConnDbOrigine As Object = ApreDB(ConnessioneDBOrigine)
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Or TypeOf (ConnDbOrigine) Is String Then
				If TypeOf (ConnGen) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
				Else
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnDbOrigine
				End If
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim idAnno As Integer = 1
				Dim NomeSquadra As String = ""

				Sql = "Select Max(idAnno)+1 From SquadraAnni Where idSquadra=" & idSquadra
				Rec = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec(0).Value Is DBNull.Value Then
						idAnno = 1
					Else
						idAnno = Rec(0).Value
					End If
				End If

				Sql = "Select Descrizione From Squadre Where idSquadra=" & idSquadra
				Rec = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Squadra non trovata"
						Ok = False
					Else
						NomeSquadra = Rec(0).Value
					End If
				End If

				Dim sAnno As String = Format(idAnno, "0000")
				Dim sCodSquadra As String = idSquadra.Trim
				While sCodSquadra.Length <> 5
					sCodSquadra = "0" & sCodSquadra
				End While
				Dim nomeDb As String = sAnno & "_" & sCodSquadra

				Sql = "Create Database [" & nomeDb & "]"
				Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
				If Ritorno.Contains(StringaErrore) Then
					Ok = False
				End If

				If Ok Then
					Sql = "Begin transaction"
					Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

					Dim Tabelle(0) As String
					Dim q As Integer = 0

					Sql = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'"
					Rec = LeggeQuery(ConnDbOrigine, Sql, ConnessioneDBOrigine)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof()
							ReDim Preserve Tabelle(q)
							Tabelle(q) = Rec("TABLE_NAME").Value

							q += 1
							Rec.MoveNext()
						Loop
						Rec.Close()
					End If

					If Ok Then
						For i As Integer = 0 To q - 1
							Try
								Sql = "Select * Into [" & nomeDb & "].[dbo].[" & Tabelle(i) & "] From " & Tabelle(i)
								Ritorno = EsegueSql(ConnDbOrigine, Sql, ConnessioneDBOrigine)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						Next

						If Ok Then
							Dim ConnessioneNuovo As String = LeggeImpostazioniDiBase(Server.MapPath("."), nomeDb)
							Dim ConnNuovo As Object = ApreDB(ConnessioneNuovo)

							If TypeOf (ConnNuovo) Is String Then
								Ritorno = ErroreConnessioneDBNonValida & ":" & ConnNuovo
								Ok = False
							Else
								For i As Integer = 0 To q - 1
									Sql = "Select Chiave From _CHIAVI_ Where Upper(lTrim(rTrim(Tabella)))='" & Tabelle(i).Trim.ToUpper & "'"
									Rec = LeggeQuery(ConnNuovo, Sql, ConnessioneNuovo)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
										Exit For
									Else
										If Not Rec.Eof() Then
											Dim Query As String = Rec(0).Value

											Ritorno = EsegueSql(ConnNuovo, Query, ConnessioneNuovo)
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
												Exit For
											End If

											Rec.Close()
										End If
									End If
								Next

								If Ok Then
									Sql = "Drop Table _CHIAVI_"
									Ritorno = EsegueSql(ConnNuovo, Sql, ConnessioneNuovo)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
								End If
							End If
						End If

						If Ok Then
							Sql = "Insert Into [" & nomeDb & "].[dbo].[Anni] Values (" &
								" " & idAnno & ", " &
								"'" & NuovoAnno & "', " &
								"'" & NomeSquadra.Replace("'", "''") & "', " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null, " &
								"null " &
								")"
							Ritorno = EsegueSql(ConnDbOrigine, Sql, ConnessioneDBOrigine)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							Else
								Sql = "Insert SquadraAnni Values (" &
									" " & idSquadra & ", " &
									" " & idAnno & ", " &
									"'" & NuovoAnno & "' " &
									")"
								Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								Else
								End If
							End If
						End If
					End If
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

					Sql = "Drop Database [" & nomeDb & "]"
					Ritorno2 = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ImportaAnagrafica(CodiceSquadra As String, Squadra As String, idAnno As String) As String
		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory
		Dim Path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
		Dim NomeFile As String = Path.Trim & Squadra.Replace(" ", "_").Trim & "\CSV\importAnagrafica.csv"
		Dim CampiCSV() As String = {"Cognome", "Nome", "EMail", "Telefono", "DataDiNascita", "Indirizzo", "CodFiscale", "Maschio", "Citta", "Cap"}
		Dim TipoCampiCSV() As String = {"T", "T", "T", "N", "T", "T", "T", "T", "T", "T"}

		If Not File.Exists(NomeFile) Then
			Ritorno = StringaErrore & " File non esistente: " & NomeFile
		Else
			Dim Tutto As String = gf.LeggeFileIntero(NomeFile)
			Dim Righe() As String = Tutto.Split(vbCrLf)

			If Righe.Count = 0 Then
				Ritorno = StringaErrore & " File vuoto"
			Else
				Dim Campi() As String = Righe(0).Split(";")

				If Campi.Count = 0 Then
					Ritorno = StringaErrore & " Intestazione vuota"
				Else
					If Campi.Count <> CampiCSV.Count Then
						Ritorno = StringaErrore & " Intestazione non valida"
					Else
						Dim q As Integer = 0

						For Each c In CampiCSV
							If c.Trim.ToUpper <> Campi(q).Trim.ToUpper Then
								Ritorno = StringaErrore & " Intestazione non valida: " & CampiCSV.ToString & " -> " & Campi.ToString
								Exit For
							End If
							q += 1
						Next

						If Ritorno = "" Then
							Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), CodiceSquadra)

							If ConnessioneGenerale = "" Then
								Ritorno = ErroreConnessioneNonValida
							Else
								Dim ConnGen As Object = ApreDB(ConnessioneGenerale)
								Dim Ok As Boolean = True
								Dim Datella As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")

								gf.ApreFileDiTestoPerScrittura(Path.Trim & Squadra.Replace(" ", "_").Trim & "\CSV\LogCaricamento_" & Datella & ".txt")
								gf.ScriveTestoSuFileAperto("Codice squadra: " & CodiceSquadra)

								If TypeOf (ConnGen) Is String Then
									Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
								Else
									gf.ScriveTestoSuFileAperto("Begin trans")

									Dim Sql As String = "Begin transaction"
									Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

									Dim Scartati As Integer = 0
									Dim Scritti As Integer = 0

									If Ritorno = "*" Then
										Try
											Dim IntestCampi As String = ""

											gf.ScriveTestoSuFileAperto("Intestazione 1")

											For i As Integer = 0 To CampiCSV.Count - 1
												IntestCampi &= CampiCSV(i) & ", "
											Next
											IntestCampi = "(idAnno, idGiocatore, idCategoria, " & Mid(IntestCampi, 1, IntestCampi.Length - 2) & ", Eliminato, RapportoCompleto, " &
												"idRuolo, Maggiorenne, idTaglia, Categorie, idCategoria2, idCategoria3 )"

											gf.ScriveTestoSuFileAperto("Intestazione 2")

											Dim idGiocatore As Integer = 1
											Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
											Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")

											Sql = "Select Max(idGiocatore)+1 From Giocatori"
											Rec = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
											If TypeOf (Rec) Is String Then
												Ritorno = Rec
												Ok = False
											Else
												If Rec(0).Value Is DBNull.Value = True Then
													idGiocatore = 1
												Else
													idGiocatore = Rec(0).Value
													Rec.Close
												End If
											End If
											gf.ScriveTestoSuFileAperto("idGiocatore di partenza: " & idGiocatore)

											gf.ScriveTestoSuFileAperto("Righe: " & Righe.Count - 1)

											For i As Integer = 1 To Righe.Count - 1
												If Righe(i).Trim <> "" Then
													Dim Scrive As Boolean = True
													Dim Campi2() As String = Righe(i).Split(";")
													Sql = "Insert Into Giocatori " & IntestCampi & " Values ("

													Sql &= idAnno & ", " & idGiocatore & ", -1, "

													' gf.ScriveTestoSuFileAperto("Riga: " & Righe(i))
													' gf.ScriveTestoSuFileAperto("Campi: " & Campi2.Count - 1)

													Dim eMail As String = ""
													Dim Maggiorenne As String = ""

													For k As Integer = 0 To Campi2.Count - 1
														Dim c As String = IIf(Campi2(k) = "", "null", Campi2(k))

														If CampiCSV(k).ToUpper.Contains("CODFISCALE") Then
															Dim Sql1 As String = "Select * From Giocatori Where Upper(Ltrim(Rtrim(CodFiscale)))='" & Campi2(k).Trim.ToUpper & "'"
															Rec2 = LeggeQuery(ConnGen, Sql1, ConnessioneGenerale)
															If TypeOf (Rec2) Is String Then
																Ritorno = Rec2
																Ok = False
															Else
																If Not Rec2.Eof Then
																	Scrive = False
																	Rec2.Close
																	Exit For
																End If
															End If
														End If

														If CampiCSV(k).ToUpper.Contains("MAIL") Then
															eMail = Campi2(k)
														End If

														If CampiCSV(k).ToUpper.Contains("MASCHIO") Then
															If Campi2(k) = "S" Or Campi2(k) = "M" Then
																Campi2(k) = "M"
															Else
																Campi2(k) = "F"
															End If
														End If

														If CampiCSV(k).ToUpper.Contains("NASCITA") Then
															Dim d() As String = Campi2(k).Split("/")
															Campi2(k) = d(2) & "-" & d(1) & "-" & d(0)
															Dim dd As Date = Convert.ToDateTime(Campi2(k))
															Dim Oggi As Date = Now
															Dim diff As Integer = DateDiff(DateInterval.Year, dd, Oggi)
															If diff >= 18 Then
																Maggiorenne = "S"
															Else
																Maggiorenne = "N"
															End If
														End If

														If TipoCampiCSV(k) = "T" Then
															Sql &= "'" & Campi2(k).Replace(vbCrLf, "").Replace("'", "''").Trim() & "', "
														Else
															Sql &= Campi2(k).Replace(vbCrLf, "").Trim() & ", "
														End If
													Next
													Sql = Mid(Sql, 1, Sql.Length - 2) & ", 'N', 'N', 0, '" & Maggiorenne & "', -1, '', -1, -1 "
													Sql &= ")"

													If Scrive = True Then
														idGiocatore += 1

														gf.ScriveTestoSuFileAperto("Maggiorenne:" & Maggiorenne)
														gf.ScriveTestoSuFileAperto("EMail:" & eMail)
														gf.ScriveTestoSuFileAperto(Sql)

														Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

														If Ritorno.Contains(StringaErrore) Then
															gf.ScriveTestoSuFileAperto(Ritorno)
															Ok = False
															Exit For
														Else
															Sql = "Insert into GiocatoriDettaglio (idAnno, idGiocatore) Values (" & idAnno & ", " & idGiocatore & ")"
															gf.ScriveTestoSuFileAperto(Sql)
															Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

															If Ritorno.Contains(StringaErrore) Then
																gf.ScriveTestoSuFileAperto(Ritorno)
																Ok = False
																Exit For
															Else
																Sql = "Insert into GiocatoriMails Values (" & idGiocatore & ", 3, '" & eMail.Replace("'", "''") & "', 'S')"
																gf.ScriveTestoSuFileAperto(Sql)
																Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

																Sql = "Insert into GiocatoriMails Values (" & idGiocatore & ", 1, '', 'S')"
																gf.ScriveTestoSuFileAperto(Sql)
																Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

																Sql = "Insert into GiocatoriMails Values (" & idGiocatore & ", 2, '', 'S')"
																gf.ScriveTestoSuFileAperto(Sql)
																Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

																Sql = "Insert Into GiocatoriSemafori Values (" &
																	" " & idGiocatore & ", " &
																	"'rosso', " &
																	"'Giocatore non iscritto', " &
																	"'rosso', " &
																	"'Pagamento non completo', " &
																	"'rosso', " &
																	"'Nessuna firma validata', " &
																	"'rosso', " &
																	"'Flag certificato non impostato', " &
																	"'rosso', " &
																	"'Nessun elemento kit consegnato' " &
																	")"
																Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

																gf.ScriveTestoSuFileAperto("Riga scritta")
																gf.ScriveTestoSuFileAperto("")

																Scritti += 1
															End If
														End If
													Else
														Scartati += 1
														gf.ScriveTestoSuFileAperto("Riga scartata.")
														gf.ScriveTestoSuFileAperto("")
													End If
													If Ritorno <> "*" Then
														Ok = False
														Exit For
													End If
												End If
											Next
										Catch ex As Exception
											gf.ScriveTestoSuFileAperto(ex.Message)
											Ritorno = StringaErrore & " " & ex.Message
										End Try

										If Ritorno = "*" Then
											Ok = True
											Ritorno = Scritti & ";" & Scartati ' Righe.Count - 3
										End If

										If Ok Then
											gf.EliminaFileFisico(NomeFile)

											Sql = "commit"
											Dim Ritorno2 As String = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
										Else
											Sql = "rollback"
											Dim Ritorno2 As String = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
										End If
									End If
								End If

								gf.ChiudeFileDiTestoDopoScrittura()
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function OnLineOffLineSquadra(idAnno As String, idSquadra As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = ApreDB(ConnessioneGenerale)
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Stato As String = ""

				' Sql = "Select * From SquadraAnni Where idAnno=" & idAnno & " And idSquadra=" & idSquadra
				Sql = "Select * From SquadraAnni Where idSquadra=" & idSquadra
				Rec = LeggeQuery(ConnGen, Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessna squadra rilevata"
						Rec.Close
					Else
						Stato = Rec("OnLine").Value
						Rec.Close

						If Stato = "S" Then
							Stato = "N"
						Else
							Stato = "S"
						End If

						Try
							Sql = "Update SquadraAnni Set " &
								"OnLine='" & Stato & "' " &
								"Where idSquadra=" & idSquadra
							' "Where idAnno=" & idAnno & " And idSquadra=" & idSquadra
							Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
							If Not Ritorno.Contains(StringaErrore) Then
								Ritorno = "*"
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
						End Try
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class