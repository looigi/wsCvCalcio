Imports System.Web.Services
Imports System.ComponentModel
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
	Public Function PulisceDBDaTrial(NomeDb As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim Ok As Boolean = True

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Sql As String = ""
			Dim ConnGen As Object = New clsGestioneDB
			Dim Rec As Object

			Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
			Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

			Sql = "SELECT Distinct TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' And TABLE_SCHEMA='" & NomeDb & "'"
			Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale, False)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
				Ok = False
			Else
				Ritorno = ""
				Do Until Rec.Eof()
					Dim Tabella As String = Rec(0).Value
					Sql = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '" & NomeDb & "' AND TABLE_NAME = '" & Tabella & "' AND COLUMN_NAME Like '%trial%'"
					Dim Rec2 As Object
					Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale, False)
					Do Until Rec2.Eof()
						Sql = "ALTER TABLE " & NomeDb & "." & Tabella & " DROP " & Rec2(0).Value & ""
						Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale, False)
						If Ritorno2.Contains(StringaErrore) Then
							Ritorno = Ritorno2
							Ok = False
							Exit Do
						Else
							Ritorno &= "Tabella " & Tabella & " Drop " & Rec2(0).Value & "§"
						End If
						Rec2.MoveNext()
					Loop
					Rec2.Close()

					If Not Ok Then
						Exit Do
					End If
					Rec.MoveNext()
				Loop
				Rec.Close()
			End If

			If Ok Then
				Sql = "Commit"
				Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
			Else
				Sql = "Rollback"
				Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
			End If
		End If

		Return Ritorno
	End Function

	' http://192.168.0.205:1010/wsSuperUser.asmx?page=op&tab=test&op=CreaDB&bnd=wsSuperUserSoap&ext=testform&Squadra=Nuova+Societa&DataScadenza=31%2F12%2F2050&MailAdmin=looigi%40gmail.com&NomeAdmin=Luigi&CognomeAdmin=Pecce&Anno=2022%2F2023&idTipologia=1&idLicenza=1&Mittente=looigi%40gmail.com&Telefono=5398435987&CAP=00132&Citta=Roma&Indirizzo=Via+delle+zucchinelle&Stima=1&PIVA=84848&CF=3939349&DBPrecompilato=N

	<WebMethod()>
	Public Function CreaDB(Squadra As String, DataScadenza As String, MailAdmin As String, NomeAdmin As String, CognomeAdmin As String, Anno As String, idTipologia As String,
						   idLicenza As String, Mittente As String, Telefono As String, CAP As String, Citta As String, Indirizzo As String, Stima As String, PIVA As String, CF As String, DBPrecompilato As String) As String
		Dim Ritorno As String = ""
		Dim NomeDBDaCopiare As String = "DBVuoto"
		Dim TipoDB2 As String = "Vuoto"
		If DBPrecompilato = "S" Or DBPrecompilato.ToUpper.Trim = "TRUE" Then
			If TipoDB = "SQLSERVER" Then
				NomeDBDaCopiare = "DBPieno"
			Else
				NomeDBDaCopiare = "dbPieno"
			End If
			TipoDB2 = "Precompilato"
		End If
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim ConnessioneDBVuoto As String = LeggeImpostazioniDiBase(Server.MapPath("."), NomeDBDaCopiare)

		Dim nomeDb As String = ""

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = New clsGestioneDB
			Dim ConnDbVuoto As Object = New clsGestioneDB
			Dim Ok As Boolean = True
			Dim BarraFile As String = "\"
			Dim BarraUrl As String = "/"
			Dim idSocieta As Integer = -1

			If TipoDB <> "SQLSERVER" Then
				BarraFile = "/"
			End If

			If TypeOf (ConnGen) Is String Or TypeOf (ConnDbVuoto) Is String Then
				If TypeOf (ConnGen) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
				Else
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnDbVuoto
				End If
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim gf As New GestioneFilesDirectory

				Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim p() As String = paths.Split(";")
				If Strings.Right(p(0), 1) <> BarraFile Then
					p(0) = p(0) & BarraFile
				End If
				p(0) = p(0).Replace(vbCrLf, "")

				If Strings.Right(p(2), 1) <> BarraUrl Then
					p(2) = p(2) & BarraUrl
				End If
				p(2) = p(2).Replace(vbCrLf, "")
				p(2) = p(2).Replace("/Multimedia", "")

				Dim pathLog As String = p(1)
				If Not pathLog.EndsWith("/") Then
					pathLog &= "/"
				End If

				Dim NomeFileLog As String = pathLog & "CreazioneSocieta_" & Squadra & "_" & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".txt"

				gf.CreaDirectoryDaPercorso(pathLog)
				gf.ApreFileDiTestoPerScrittura(NomeFileLog)
				gf.ScriveTestoSuFileAperto("Squadra: " & Squadra)
				gf.ScriveTestoSuFileAperto("DataScadenza: " & DataScadenza)
				gf.ScriveTestoSuFileAperto("MailAdmin: " & MailAdmin)
				gf.ScriveTestoSuFileAperto("NomeAdmin: " & NomeAdmin)
				gf.ScriveTestoSuFileAperto("CognomeAdmin: " & CognomeAdmin)
				gf.ScriveTestoSuFileAperto("Anno: " & Anno)
				gf.ScriveTestoSuFileAperto("idTipologia: " & idTipologia)
				gf.ScriveTestoSuFileAperto("idLicenza: " & idLicenza)
				gf.ScriveTestoSuFileAperto("Mittente: " & Mittente)
				gf.ScriveTestoSuFileAperto("Tipologia DB: " & TipoDB2)
				gf.ScriveTestoSuFileAperto("Server DB: " & TipoDB)
				gf.ScriveTestoSuFileAperto("-------------------------------------------")

				Sql = "Select * From DettaglioSocieta Where EMailAmministratore='" & MailAdmin.Replace("'", "''") & "'"
				Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Not Rec.Eof() Then
						gf.ScriveTestoSuFileAperto("Mail Admin già presente in archivio: " & MailAdmin)
						Return StringaErrore & " Mail Admin già presente in archivio"
					End If
					Rec.Close()
				End If

				If TipoDB = "SQLSERVER" Then
					Sql = "Select IsNull(Max(idSquadra),0)+1 From Squadre"
				Else
					Sql = "Select Coalesce(Max(idSquadra),0)+1 From Squadre"
				End If
				Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					'If Rec(0).Value Is DBNull.Value = True Then
					'	idSocieta = 1
					'Else
					idSocieta = Rec(0).Value
					Rec.Close()
					'End If
					gf.ScriveTestoSuFileAperto("idSocieta: " & idSocieta.ToString)

					If TipoDB = "SQLSERVER" Then
						Sql = "Select IsNull(Max(idUtente),0)+1 From Utenti"
					Else
						Sql = "Select Coalesce(Max(idUtente),0)+1 From Utenti"
					End If
					Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Dim idUtente As Integer = -1

						'If Rec(0).Value Is DBNull.Value = True Then
						'	idUtente = 1
						'Else
						idUtente = Rec(0).Value
						Rec.Close()
						'End If
						gf.ScriveTestoSuFileAperto("idUtente: " & idUtente.ToString)

						If Ok Then
							Dim Tabelle(0) As String
							Dim q As Integer = 0

							Sql = "SELECT Distinct TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' And TABLE_SCHEMA='" & NomeDBDaCopiare & "'"
							Rec = ConnDbVuoto.LeggeQuery(Server.MapPath("."), Sql, ConnessioneDBVuoto, False)
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

							'Return "Tabelle da copiare: " & Tabelle.Length - 1 & " -> " & Ritorno

							If Ok Then
								Dim Societa As String = Format(idSocieta, "00000")
								nomeDb = "0001_" & Societa

								gf.ScriveTestoSuFileAperto("idSocieta 2: " & Societa)
								gf.ScriveTestoSuFileAperto("Nome DB: " & nomeDb)

								Sql = "Create Database [" & nomeDb & "]"
								Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
									gf.ScriveTestoSuFileAperto(Ritorno & " -> " & Sql)
								Else
									gf.ScriveTestoSuFileAperto("DB Creato")
								End If

								If Ok Then
									Dim ConnessioneNuovo As String = LeggeImpostazioniDiBase(Server.MapPath("."), nomeDb)
									Dim ConnNuovo As Object = New clsGestioneDB

									Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
									Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

									Sql = "Insert Into Squadre Values (" & idSocieta & ", '" & Squadra.Replace("'", "''") & "', '" & DataScadenza & "', " & idTipologia & ", " & idLicenza & ", 'N', " & Now.Month & ", " & Now.Year & ")"
									Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
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
										Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
											gf.ScriveTestoSuFileAperto(Ritorno & " -> " & Sql)
										Else
											gf.ScriveTestoSuFileAperto("Inserita riga tabella Squadre")
										End If

										If Ok Then
											gf.ScriveTestoSuFileAperto("Utente Inserito")

											For i As Integer = 0 To q - 1
												gf.ScriveTestoSuFileAperto("Copia Tabella: " & Tabelle(i))
												Try
													If TipoDB = "SQLSERVER" Then
														Sql = "Select * Into [" & nomeDb & "].[dbo].[" & Tabelle(i) & "] From " & Tabelle(i)
													Else
														Sql = "CREATE TABLE " & nomeDb & "." & Tabelle(i) & " SELECT * FROM " & Tabelle(i)
													End If
													Ritorno = ConnDbVuoto.EsegueSql(Server.MapPath("."), Sql, ConnessioneDBVuoto)
													If Ritorno.Contains(StringaErrore) Then
														Ok = False
														gf.ScriveTestoSuFileAperto("Errore su copia Tabella: " & Tabelle(i) & " -> " & Ritorno)
														Exit For
													Else
														gf.ScriveTestoSuFileAperto("Copiata Tabella: " & Tabelle(i) & " (" & Sql & ")")
													End If
												Catch ex As Exception
													gf.ScriveTestoSuFileAperto("ERRORE su Copia Tabella: " & Tabelle(i) & " -> " & ex.Message)
													Ritorno = StringaErrore & " " & ex.Message
													Ok = False
												End Try
											Next

											If Ok Then
												If TypeOf (ConnNuovo) Is String Then
													Ritorno = ErroreConnessioneDBNonValida & ":" & ConnNuovo
													Ok = False
												Else
													For i As Integer = 0 To q - 1
														gf.ScriveTestoSuFileAperto("Gestione chiavi Tabella: " & Tabelle(i))

														Sql = "Select Chiave From _CHIAVI_ Where Upper(lTrim(rTrim(Tabella)))='" & Tabelle(i).Trim.ToUpper & "'"
														Rec = ConnDbVuoto.LeggeQuery(Server.MapPath("."), Sql, ConnessioneDBVuoto)
														If TypeOf (Rec) Is String Then
															gf.ScriveTestoSuFileAperto("ERRORE creazione recordset Gestione chiavi Tabella: " & Tabelle(i) & " -> " & Rec)
															Ritorno = Rec
															Ok = False
															Exit For
														Else
															If Rec.Eof() = False Then
																Dim Query As String = "" & Rec("Chiave").Value

																If Query <> "" Then
																	gf.ScriveTestoSuFileAperto("Chiave Tabella: " & Tabelle(i) & " -> " & Query)

																	If TipoDB <> "SQLSERVER" Then
																		Query = Query.ToLower
																		'Query = Mid(Query, 1, Query.ToLower.IndexOf("with"))
																	End If
																	Ritorno = ConnNuovo.EsegueSql(Server.MapPath("."), Query, ConnessioneNuovo)
																	If Ritorno.Contains(StringaErrore) Then
																		gf.ScriveTestoSuFileAperto("Errore su creazione Chiave: " & Tabelle(i) & " -> " & Ritorno)
																		Ok = False
																		Exit For
																	Else
																		gf.ScriveTestoSuFileAperto("Creata Chiave: " & Tabelle(i) & " (" & Query & ")")
																	End If
																End If

																'Rec.Close()
															End If
														End If
													Next

													If Ok Then
														gf.ScriveTestoSuFileAperto("Eliminazione Tabella Chiave")

														Sql = "Drop Table _CHIAVI_"
														Ritorno = ConnNuovo.EsegueSql(Server.MapPath("."), Sql, ConnessioneNuovo)
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
												gf.ScriveTestoSuFileAperto("Inserimento riga su tabella anni")

												Sql = "Insert Into [" & nomeDb & "].[dbo].[Anni] Values (" &
														"1, " & ' idAnno
														"'" & Anno & "', " & ' Descrizione
														"'', " & ' NomeSquadra
														"null, " & ' Lat
														"null, " & ' Lon
														"'" & Indirizzo.Replace("'", "''") & "', " &
														"'Campo " & Squadra.Replace("'", "''") & "', " & ' CampoSquadra
														"'" & Squadra.Replace("'", "''") & "', " & ' NomePolisportiva
														"'" & MailAdmin.Replace("'", "''") & "', " &
														"null, " & ' PEC
														"'" & Telefono.Replace("'", "''") & "', " &
														"'" & PIVA.Replace("'", "''") & "', " &
														"'" & CF.Replace("'", "''") & "', " &
														"null, " & ' CodiceUnivoco
														"null, " & ' SitoWeb
														"'" & MailAdmin.Replace("'", "''") & "', " & ' MittenteMail
														"null, " & ' GestionePagamenti
														"null, " & ' CostoScuolaCalcio
														"null, " & ' Suffisso
														"null, " & ' iscrFirmaEntrambi
														"null, " & ' Modulo Associato
														"10, " & ' PercCashBack
														"'N', " & ' Rate Manuali
														"'N' " & ' Cashback
														")"
												Ritorno = ConnDbVuoto.EsegueSql(Server.MapPath("."), Sql, ConnessioneDBVuoto)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
													gf.ScriveTestoSuFileAperto("Creazione voce su Tabella Anni: " & Ritorno)
												Else
													gf.ScriveTestoSuFileAperto("Dati inseriti in tabella Anni: " & "[" & nomeDb & "].[dbo].[Anni]")
													gf.ScriveTestoSuFileAperto("Insertimento riga SquadraAnni")

													Sql = "Insert Into SquadraAnni Values (" &
														" " & idSocieta & ", " &
														"1, " &
														"'" & Anno & "', " &
														"'S' " &
														")"
													Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
													If Ritorno.Contains(StringaErrore) Then
														Ok = False
														gf.ScriveTestoSuFileAperto(Ritorno & " -> " & Sql)
													Else
														gf.ScriveTestoSuFileAperto("Dati inseriti in tabella SquadraAnni")

														'Crea riga su dettaglio società
														Sql = "Insert Into [Generale].[dbo].[DettaglioSocieta] Values (" &
															"'" & nomeDb & "', " &
															"'" & Squadra.Replace("'", "''") & "', " &
															"'" & MailAdmin.Replace("'", "''") & "', " &
															"'" & NomeAdmin.Replace("'", "''") & "', " &
															"'" & CognomeAdmin.Replace("'", "''") & "', " &
															"'" & Telefono.Replace("'", "''") & "', " &
															"'" & CAP.Replace("'", "''") & "', " &
															"'" & Citta.Replace("'", "''") & "', " &
															"'" & Indirizzo.Replace("'", "''") & "', " &
															" " & Stima & ", " &
															"'" & PIVA.Replace("'", "''") & "', " &
															"'" & CF.Replace("'", "''") & "', " &
															" " & idLicenza & ", " &
															" " & idTipologia & " " &
															")"
														'Crea riga su dettaglio società
														Ritorno = ConnDbVuoto.EsegueSql(Server.MapPath("."), Sql, ConnessioneDBVuoto)
														If Ritorno.Contains(StringaErrore) Then
															Ok = False
															gf.ScriveTestoSuFileAperto("Creazione voce su Dettaglio Società: " & Ritorno & vbCrLf & Sql)
														Else
															gf.ScriveTestoSuFileAperto("Dati inseriti in tabella Dettaglio società: " & "[" & nomeDb & "].[dbo].[Anni]")

															'If TipoDB <> "SQLSERVER" Then
															'	gf.ScriveTestoSuFileAperto("Concessione privilegi all'utente sul db della nuova società")

															'	Dim ConnessioneGeneraleAdmin As String = LeggeImpostazioniDiBase(Server.MapPath("."), "", True)
															'	Dim ConnGenAdmin As Object = New clsGestioneDB

															'	Sql = "GRANT ALL PRIVILEGES ON `" & nomeDb & "`.* TO 'incalciouser'@'%' WITH GRANT OPTION;"
															'	Ritorno = ConnGenAdmin.EsegueSql(Server.MapPath("."), Sql, ConnessioneGeneraleAdmin, False)
															'	If Ritorno.Contains(StringaErrore) Then
															'		Ok = False
															'		gf.ScriveTestoSuFileAperto("Errore nell'impostazione dei permessi sul db della nuova società: " & Ritorno)
															'	Else
															'		gf.ScriveTestoSuFileAperto("Impostati i permessi sul db della nuova società")
															'	End If
															'End If

															If Ok Then
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
																	Dim BodyMail As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\template_nuova_societa\template-mail-nuova-societa.html")

																	Dim Body As String = ""
																	Body &= Squadra & "<br /><br />"
																	Body &= "Amministratore: " & CognomeAdmin & " " & NomeAdmin & "<br />"
																	Body &= "Scadenza licenza: " & DataScadenza & "<br />"
																	Body &= "Anno: " & Anno & "<br />"
																	Body &= "Tipologia: " & Tipologia & "<br />"
																	Body &= "Licenza: " & Licenza & "<br /><br />"
																	Body &= "Per accedere, l'amministratore potrà utilizzare la password " & nuovaPass(0) & " che dovrà essere modificata al primo ingresso."

																	BodyMail = BodyMail.Replace("***TESTO MAIL***", Body)

																	Dim urlIMG As String = p(2) & "Scheletri\template_nuova_societa\images\LOGOinCalcio200n.png"
																	Dim contentFB As String = p(2) & "Scheletri\template_nuova_societa\images\facebook2x.png"
																	Dim contentLogo As String = p(2) & "Scheletri\template_nuova_societa\images\LOGOinCalcio200n.png"
																	Dim pathIMG As String = p(2) & "Scheletri\template_nuova_societa\images\Portatile_homeapp_1.png"

																	BodyMail = BodyMail.Replace("***URL IMG***", pathIMG)
																	BodyMail = BodyMail.Replace("***contentFB***", contentFB)
																	BodyMail = BodyMail.Replace("***contentLOGO***", contentLogo)
																	BodyMail = BodyMail.Replace("***PATH_IMG***", pathIMG)

																	Dim ChiScrive As String = "servizioclienti@incalcio.cloud"

																	gf.ScriveTestoSuFileAperto("Invio Mail")

																	Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, BodyMail, MailAdmin, {""}, "NUOVA SOCIETA")
																	If Ritorno.Contains(StringaErrore) Then
																		gf.ScriveTestoSuFileAperto("Ritorno invio mail destinario " & MailAdmin & ": " & Ritorno)
																	Else
																		Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, Body, Mittente, {""})
																		If Ritorno.Contains(StringaErrore) Then
																			gf.ScriveTestoSuFileAperto("Ritorno invio mail destinario " & Mittente & ": " & Ritorno)
																		Else
																			Ritorno = Societa

																			If TipoDB = "SQLSERVER" Then
																				' Copia immagine di base
																				Dim pathImmagini As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
																				pathImmagini = pathImmagini.Replace(vbCrLf, "")
																				If Strings.Right(pathImmagini, 1) <> BarraFile Then
																					pathImmagini &= BarraFile
																				End If
																				Dim Dest1 As String = p(0) & nomeDb & "\Societa\Societa_1.png"
																				Dim Dest2 As String = p(0) & nomeDb & "\Societa\Societa_2.png"
																				gf.CreaDirectoryDaPercorso(Dest1)

																				If TipoDB <> "SQLSERVER" Then
																					pathImmagini = pathImmagini.Replace("\", "/")
																					pathImmagini = pathImmagini.Replace("//", "/")

																					Dest1 = Dest1.Replace("\", "/")
																					Dest1 = Dest1.Replace("//", "/")

																					Dest2 = Dest2.Replace("\", "/")
																					Dest2 = Dest2.Replace("//", "/")
																				End If

																				Try
																					gf.ScriveTestoSuFileAperto("Copia immagini societa: " & pathImmagini & "Sconosciuto.png" & " -> " & Dest1)
																					File.Copy(pathImmagini & "Sconosciuto.png", Dest1)

																					gf.ScriveTestoSuFileAperto("Copia immagini societa: " & pathImmagini & "Sconosciuto.png" & " -> " & Dest2)
																					File.Copy(pathImmagini & "Sconosciuto.png", Dest2)
																				Catch ex As Exception

																				End Try
																			End If
																		End If
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
					End If
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Commit: " & Ritorno2)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Rollback: " & Ritorno2)

					ConnDbVuoto.Close()

					Sql = "Drop Database [" & nomeDb & "]"
					Ritorno2 = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Drop Database: " & Ritorno2)

					Sql = "Delete From [Generale].[dbo].[Squadre] Where idsquadra='" & idSocieta & "'"
					Ritorno2 = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Delete from DettaglioSocieta: " & Ritorno2)

					Sql = "Delete From [Generale].[dbo].[SquadraAnni] Where idsquadra='" & idSocieta & "'"
					Ritorno2 = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Delete from DettaglioSocieta: " & Ritorno2)

					Sql = "Delete From [Generale].[dbo].[Utenti] Where idsquadra='" & idSocieta & "'"
					Ritorno2 = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Delete from DettaglioSocieta: " & Ritorno2)

					Sql = "Delete From [Generale].[dbo].[DettaglioSocieta] Where codsquadra='" & nomeDb & "'"
					Ritorno2 = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					gf.ScriveTestoSuFileAperto("Delete from DettaglioSocieta: " & Ritorno2)

					ConnGen.Close()
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
			Dim ConnGen As Object = New clsGestioneDB
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object
				Dim Rec2 As Object
				Dim Sql As String = ""

				Sql = "Select * From Squadre Where Eliminata = 'N' Order By Descrizione"
				Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)

				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
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

							Dim Scadenza As String = ""
							Try
								Scadenza = "" & Rec("DataScadenza").Value
							Catch ex As Exception

							End Try

							Dim Semaforo1 As String = "" : Dim Titolo1 As String = ""
							Dim Semaforo2 As String = "" : Dim Titolo2 As String = ""

							If Scadenza <> "" Then
								Dim sc() As String = Scadenza.Split("-")
								Scadenza = sc(2) & "/" & sc(1) & "/" & sc(0)

								Dim dScadenza As DateTime
								Dim Oggi As Date = Now
								Dim diff As Integer = 0

								Try
									dScadenza = Convert.ToDateTime(Scadenza)
									diff = DateAndTime.DateDiff(DateInterval.Day, Oggi, dScadenza, )
								Catch ex As Exception

								End Try

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

							Sql = "Select " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " From SquadraAnni Where idSquadra = " & Rec("idSquadra").Value
							Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2
							Else
								Anni = Rec2(0).Value
								'End If
								Rec2.Close()
							End If

							Sql = "Select " & IIf(TipoDB = "SQLSERVER", "Top 1", "") & " * From SquadraAnni Where idSquadra = " & Rec("idSquadra").Value & " Order By idAnno Desc" & IIf(TipoDB = "SQLSERVER", "", " Limit 1")
							Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2
							Else
								If Not Rec2.Eof() Then
									maxAnno = Rec2("idAnno").Value
									Stato = Rec2("OnLine").Value
								End If
								Rec2.Close()
							End If

							If Stato = "S" Then
								Semaforo2 = "verde" : Titolo2 = "Database in linea"
							Else
								Semaforo2 = "rosso" : Titolo2 = "Database fuori linea"
							End If

							Dim id As String = Rec("idSquadra").Value.ToString.Trim
							For i As Integer = id.Length To 4
								id = "0" & id
							Next
							For i As Integer = maxAnno.Length To 3
								maxAnno = "0" & maxAnno
							Next
							Dim CodiceSquadra As String = maxAnno & "_" & id
							Dim RateManuali As String = "N"
							Dim Cashback As String = "N"
							Dim GestioneGenitori As String = "N"

							Sql = "Select RateManuali, Cashback From [" & CodiceSquadra & "].[dbo].[Anni]"
							Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2
							Else
								If Not Rec2.Eof() Then
									RateManuali = "" & Rec2("RateManuali").Value
									Cashback = "" & Rec2("Cashback").Value
								End If
								Rec2.Close()
							End If

							Sql = "Select * From GestioneGenitori Where idSquadra = " & Val(id)
							Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2
							Else
								If Not Rec2.Eof() Then
									GestioneGenitori = "" & Rec2("GestioneGenitori").Value
								End If
								Rec2.Close()
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
									RateManuali & ";" &
									Cashback & ";" &
									GestioneGenitori & ";" &
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
	Public Function ModificaSquadra(idSquadra As String, Squadra As String, DataScadenza As String, idTipologia As String, idLicenza As String, rateManuali As String, Cashback As String, GestioneGenitori As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If ConnessioneGenerale = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim ConnGen As Object = New clsGestioneDB
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					Sql = "Update Squadre Set " &
						"Descrizione='" & Squadra.Replace("'", "''") & "'," &
						"DataScadenza='" & DataScadenza & "'," &
						"idTipologia=" & idTipologia & "," &
						"idLicenza=" & idLicenza & " " &
						"Where idSquadra=" & idSquadra
					Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					If Not Ritorno.Contains(StringaErrore) Then
						Dim maxAnno As String = ""
						Sql = "Select Max(idAnno) From SquadraAnni Where idSquadra=" & idSquadra
						Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Not Rec.Eof() Then
								maxAnno = Rec(0).Value
							End If
							Rec.Close()

							Dim id As String = idSquadra.ToString.Trim
							For i As Integer = id.Length To 4
								id = "0" & id
							Next
							For i As Integer = maxAnno.Length To 3
								maxAnno = "0" & maxAnno
							Next
							Dim CodiceSquadra As String = maxAnno & "_" & id

							Sql = "Update [" & CodiceSquadra & "].[dbo].[Anni] Set rateManuali = '" & rateManuali & "', Cashback = '" & Cashback & "'"
							Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
							If Not Ritorno.Contains(StringaErrore) Then
								Ritorno = "*"
							End If

							Sql = "Select * From GestioneGenitori Where idSquadra=" & idSquadra
							Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Not Rec.Eof() Then
									Sql = "Update GestioneGenitori Set GestioneGenitori = '" & GestioneGenitori & "' Where idSquadra = " & idSquadra
								Else
									Sql = "Insert Into GestioneGenitori Values( " & idSquadra & ", '" & GestioneGenitori & "')"
								End If
								Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
								If Not Ritorno.Contains(StringaErrore) Then
									Ritorno = "*"
								End If
							End If
						End If
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
			Dim ConnGen As Object = New clsGestioneDB
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					Sql = "Update Squadre Set " &
						"Eliminata='S'" &
						"Where idSquadra=" & idSquadra
					Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
					If Not Ritorno.Contains(StringaErrore) Then
						Sql = "Update Utenti Set " &
							"Eliminato='S'" &
							"Where idSquadra=" & idSquadra
						Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
						If Not Ritorno.Contains(StringaErrore) Then
							Sql = "Select * From Squadre Where idSquadra=" & idSquadra
							Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Not Rec.Eof() Then
									Dim CodSquadra As String = Rec("CodSquadra").Value

									Sql = "Delete From dettagliosocieta " &
										"Where codsquadra='" & CodSquadra & "'"
									Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
									If Not Ritorno.Contains(StringaErrore) Then
										Sql = "Update squadraanni set Eliminata='S' " &
											"Where idsquadra=" & idSquadra
										Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
										If Not Ritorno.Contains(StringaErrore) Then
											Ritorno = "*"
										End If
									End If
								End If
							End If

						End If
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
			Dim ConnGen As Object = New clsGestioneDB
			Dim ConnDbOrigine As Object = New clsGestioneDB
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Or TypeOf (ConnDbOrigine) Is String Then
				If TypeOf (ConnGen) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
				Else
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnDbOrigine
				End If
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim idAnno As Integer = 1
				Dim NomeSquadra As String = ""

				If TipoDB = "SQLSERVER" Then
					Sql = "Select IsNull(Max(idAnno),0)+1 From SquadraAnni Where idSquadra=" & idSquadra
				Else
					Sql = "Select Coalesce(Max(idAnno),0)+1 From SquadraAnni Where idSquadra=" & idSquadra
				End If
				Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	idAnno = 1
					'Else
					idAnno = Rec(0).Value
					'End If
				End If

				Sql = "Select Descrizione From Squadre Where idSquadra=" & idSquadra
				Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
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
				Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
				If Ritorno.Contains(StringaErrore) Then
					Ok = False
				End If

				If Ok Then
					Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
					Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

					Dim Tabelle(0) As String
					Dim q As Integer = 0

					Sql = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'"
					Rec = ConnDbOrigine.LeggeQuery(Server.MapPath("."), Sql, ConnessioneDBOrigine)
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
								Ritorno = ConnDbOrigine.EsegueSql(Server.MapPath("."), Sql, ConnessioneDBOrigine)
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
							Dim ConnNuovo As Object = New clsGestioneDB

							If TypeOf (ConnNuovo) Is String Then
								Ritorno = ErroreConnessioneDBNonValida & ":" & ConnNuovo
								Ok = False
							Else
								For i As Integer = 0 To q - 1
									Sql = "Select Chiave From _CHIAVI_ Where Upper(lTrim(rTrim(Tabella)))='" & Tabelle(i).Trim.ToUpper & "'"
									Rec = ConnNuovo.LeggeQuery(Server.MapPath("."), Sql, ConnessioneNuovo)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
										Exit For
									Else
										If Not Rec.Eof() Then
											Dim Query As String = Rec(0).Value

											Ritorno = ConnNuovo.EsegueSql(Server.MapPath("."), Query, ConnessioneNuovo)
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
									Ritorno = ConnNuovo.EsegueSql(Server.MapPath("."), Sql, ConnessioneNuovo)
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
								"null, " & ' Lat
								"null, " & ' Lon
								"'', " &
								"'Campo " & Squadra.Replace("'", "''") & "', " & ' CampoSquadra
								"'" & Squadra.Replace("'", "''") & "', " & ' NomePolisportiva
								"'', " &
								"null, " & ' PEC
								"'', " &
								"'', " &
								"'', " &
								"null, " & ' CodiceUnivoco
								"null, " & ' SitoWeb
								"'', " & ' MittenteMail
								"null, " & ' GestionePagamenti
								"null, " & ' CostoScuolaCalcio
								"null, " & ' Suffisso
								"null, " & ' iscrFirmaEntrambi
								"null, " & ' Modulo Associato
								"10, " & ' PercCashBack
								"'N', " & ' Rate Manuali
								"'N' " & ' Cashback
								")"
							Ritorno = ConnDbOrigine.EsegueSql(Server.MapPath("."), Sql, ConnessioneDBOrigine)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							Else
								Sql = "Insert SquadraAnni Values (" &
									" " & idSquadra & ", " &
									" " & idAnno & ", " &
									"'" & NuovoAnno & "' " &
									")"
								Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
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
					Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

					Sql = "Drop Database [" & nomeDb & "]"
					Ritorno2 = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
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

		If Not ControllaEsistenzaFile(NomeFile) Then
			Ritorno = StringaErrore & " File non esistente: " & NomeFile
		Else
			Dim Tutto As String = gf.LeggeFileIntero(NomeFile)
			Dim Righe() As String = Tutto.Split(vbCrLf)

			If Righe.Count = 0 Then
				Ritorno = StringaErrore & " File vuoto"
			Else
				Dim Campi() As String = (Righe(0).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "")).Split(";")

				If Campi.Count = 0 Then
					Ritorno = StringaErrore & " Intestazione vuota"
				Else
					If Campi.Count <> CampiCSV.Count Then
						Ritorno = StringaErrore & " Intestazione non valida"
					Else
						Dim q As Integer = 0

						For Each c In CampiCSV
							If c.Trim.ToUpper.Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "") <> Campi(q).Trim.ToUpper.Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "") Then
								Ritorno = StringaErrore & " Intestazione non valida: " & c.ToString & " -> " & Campi(q).ToString
								Exit For
							End If
							q += 1
						Next

						If Ritorno = "" Then
							Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), CodiceSquadra)

							If ConnessioneGenerale = "" Then
								Ritorno = ErroreConnessioneNonValida
							Else
								Dim ConnGen As Object = New clsGestioneDB
								Dim Ok As Boolean = True
								Dim Datella As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")

								gf.ApreFileDiTestoPerScrittura(Path.Trim & Squadra.Replace(" ", "_").Trim & "\CSV\LogCaricamento_" & Datella & ".txt")
								gf.ScriveTestoSuFileAperto("Codice squadra: " & CodiceSquadra)

								If TypeOf (ConnGen) Is String Then
									Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
								Else
									gf.ScriveTestoSuFileAperto("Begin trans")

									Dim Sql As String = "Begin transaction"
									Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

									Dim Scartati As Integer = 0
									Dim Scritti As Integer = 0

									If Ritorno = "*" Then
										Try
											Dim IntestCampi As String = ""

											gf.ScriveTestoSuFileAperto("Intestazione 1")

											For i As Integer = 0 To CampiCSV.Count - 1
												IntestCampi &= CampiCSV(i).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "") & ", "
											Next
											IntestCampi = "(idAnno, idGiocatore, idCategoria, " & Mid(IntestCampi, 1, IntestCampi.Length - 2) & ", Eliminato, RapportoCompleto, " &
												"idRuolo, Maggiorenne, idTaglia, Categorie, idCategoria2, idCategoria3 )"

											gf.ScriveTestoSuFileAperto("Intestazione 2")

											Dim idGiocatore As Integer = 1
											Dim Rec As Object
											Dim Rec2 As Object

											If TipoDB = "SQLSERVER" Then
												Sql = "Select IsNull(Max(idGiocatore),0)+1 From Giocatori"
											Else
												Sql = "Select Coalesce(Max(idGiocatore),0)+1 From Giocatori"
											End If
											Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
											If TypeOf (Rec) Is String Then
												Ritorno = Rec
												Ok = False
											Else
												'If Rec(0).Value Is DBNull.Value = True Then
												'	idGiocatore = 1
												'Else
												idGiocatore = Rec(0).Value
												Rec.Close()
												'End If
											End If
											gf.ScriveTestoSuFileAperto("idGiocatore di partenza: " & idGiocatore)

											gf.ScriveTestoSuFileAperto("Righe: " & Righe.Count - 1)

											For i As Integer = 1 To Righe.Count - 1
												If Righe(i).Trim <> "" Then
													Dim Scrive As Boolean = True
													Dim Campi2() As String = (Righe(i).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "")).Split(";")
													Sql = "Insert Into Giocatori " & IntestCampi & " Values ("

													Sql &= idAnno & ", " & idGiocatore & ", -1, "

													' gf.ScriveTestoSuFileAperto("Riga: " & Righe(i))
													' gf.ScriveTestoSuFileAperto("Campi: " & Campi2.Count - 1)

													Dim eMail As String = ""
													Dim Maggiorenne As String = ""

													For k As Integer = 0 To Campi2.Count - 1
														Dim c As String = IIf(Campi2(k).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "") = "", "null", Campi2(k))

														If CampiCSV(k).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "").ToUpper.Contains("CODFISCALE") Then
															Dim Sql1 As String = "Select * From Giocatori Where Upper(Ltrim(Rtrim(CodFiscale)))='" & Campi2(k).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "").Trim.ToUpper & "'"
															Rec2 = ConnGen.LeggeQuery(Server.MapPath("."), Sql1, ConnessioneGenerale)
															If TypeOf (Rec2) Is String Then
																Ritorno = Rec2
																Ok = False
															Else
																If Not Rec2.Eof() Then
																	Scrive = False
																	Rec2.Close()
																	Exit For
																End If
															End If
														End If

														If CampiCSV(k).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "").ToUpper.Contains("MAIL") Then
															eMail = Campi2(k)
														End If

														If CampiCSV(k).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "").ToUpper.Contains("MASCHIO") Then
															If Campi2(k) = "S" Or Campi2(k) = "M" Then
																Campi2(k) = "M"
															Else
																Campi2(k) = "F"
															End If
														End If

														If CampiCSV(k).Replace(Chr(34), "").Replace("'", "").Replace(vbCrLf, "").ToUpper.Contains("NASCITA") Then
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

														Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

														If Ritorno.Contains(StringaErrore) Then
															gf.ScriveTestoSuFileAperto(Ritorno)
															Ok = False
															Exit For
														Else
															Sql = "Insert into GiocatoriDettaglio (idAnno, idGiocatore) Values (" & idAnno & ", " & idGiocatore & ")"
															gf.ScriveTestoSuFileAperto(Sql)
															Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

															If Ritorno.Contains(StringaErrore) Then
																gf.ScriveTestoSuFileAperto(Ritorno)
																Ok = False
																Exit For
															Else
																Sql = "Insert into GiocatoriMails Values (" & idGiocatore & ", 3, '" & eMail.Replace("'", "''") & "', 'S')"
																gf.ScriveTestoSuFileAperto(Sql)
																Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

																Sql = "Insert into GiocatoriMails Values (" & idGiocatore & ", 1, '', 'S')"
																gf.ScriveTestoSuFileAperto(Sql)
																Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

																Sql = "Insert into GiocatoriMails Values (" & idGiocatore & ", 2, '', 'S')"
																gf.ScriveTestoSuFileAperto(Sql)
																Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

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
																Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

																gf.ScriveTestoSuFileAperto("Creazione tessera NFC per il giocatore")
																Dim Ritorno2 As String = CreaNumeroTesseraNFC(Server.MapPath("."), ConnGen, ConnessioneGenerale, Squadra, idGiocatore)
																gf.ScriveTestoSuFileAperto("Creazione tessera NFC per il giocatore. Numero tessera: " & Ritorno2)

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
													If Ritorno<> "OK" Then
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
											Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)

											Dim wsWidget As New wsWidget
											wsWidget.CreaConteggi(Squadra)
											wsWidget.CreaFirmeDaValidare(Squadra, "S")
											wsWidget.CreaIndicatori(Squadra)
											wsWidget.CreaIscritti(Squadra)
											wsWidget.CreaQuoteNonSaldate(Squadra)
										Else
											Sql = "rollback"
											Dim Ritorno2 As String = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
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
			Dim ConnGen As Object = New clsGestioneDB
			Dim Ok As Boolean = True

			If TypeOf (ConnGen) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Stato As String = ""

				' Sql = "Select * From SquadraAnni Where idAnno=" & idAnno & " And idSquadra=" & idSquadra
				Sql = "Select * From SquadraAnni Where idSquadra=" & idSquadra
				Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessna squadra rilevata"
						Rec.Close()
					Else
						Stato = Rec("OnLine").Value
						Rec.Close()

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
							Ritorno = ConnGen.EsegueSql(Server.MapPath("."), Sql, ConnessioneGenerale)
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