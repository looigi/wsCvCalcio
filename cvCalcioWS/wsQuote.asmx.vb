Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://quote.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsQuote
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function StampaListaRicevute(Squadra As String, NomeSquadra As String, DataInizio As String, DataFine As String) As String
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

				Sql = "Select A.idGiocatore, Progressivo, Pagamento, DataPagamento, B.Cognome, B.Nome, A.Validato, Case A.idTipoPagamento When 1 Then 'Rata' When 2 Then 'Altro' Else '' End As TipoPagamento, " &
						"A.idRata, A.Note, A.idUtentePagatore, A.Commento, B.Maggiorenne, A.NumeroRicevuta, C.MetodoPagamento, D.Cognome + ' ' + D.Nome As Nominativo, A.idTipoPagamento From " &
						"GiocatoriPagamenti A " &
						"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
						"Left Join MetodiPagamento C On A.MetodoPagamento = C.idMetodoPagamento " &
						"Left Join [Generale].[dbo].[Utenti] D On A.idUtenteRegistratore = D.idUtente " &
						"Where A.Eliminato = 'N' And DataPagamento Between '" & DataInizio & "' And '" & DataFine & "' " &
						"Order By A.NumeroRicevuta Desc"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna ricevuta rilevata"
					Else
						Ritorno = "<table style=""width: 100%;"" cellpadding=""0"" cellspacing=""0"">"
						Ritorno &= "<tr><th>Numero Ricevuta</th><th>Data Pagamento</th><th>Importo</th><th>Nominativo</th><th>Stato</th><th>Tipo Pagamento</th><th>Metodo Pagamento</th><th>Registrante</th><th>Descrizione</th><th>Note</th></tr>"
						Dim totale As Single = 0
						Dim totRata As Single = 0
						Dim totAltro As String = 0
						Dim metodi As New List(Of String)
						Dim totMetodi As New List(Of Single)

						Do Until Rec.Eof
							Dim Stato As String = ""
							Dim TipoPag As String = ""

							If ("" & Rec("Validato").Value) = "S" Then
								Stato = "Validata"
							Else
								Stato = "Bozza"
							End If

							Dim pag As String = ("" & Rec("Pagamento").Value) '.replace(",", ".")
							Dim pag2 As String = FormatCurrency(pag)
							Dim d As String = Rec("DataPagamento").Value
							Dim sData As String = ""
							If d.Contains("-") Then
								Dim dd() As String = d.Split("-")
								sData = dd(2) & "/" & dd(1) & "/" & dd(0)
							End If
							Ritorno &= "<tr>"
							Ritorno &= "<td style=""padding-left: 3px; width: 5%;"">" & Rec("NumeroRicevuta").Value & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 5%;"">" & sData & "</td>"
							Ritorno &= "<td style=""text-align: right; padding-left: 3px; width: 10%;"">" & pag2 & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 15%;"">" & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 7%;"">" & Stato & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 9%;"">" & Rec("TipoPagamento").Value & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 9%;"">" & Rec("MetodoPagamento").Value & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 9%;"">" & Rec("Nominativo").Value & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 15%;"">" & Rec("Commento").Value & "</td>"
							Ritorno &= "<td style=""padding-left: 3px; width: 15%;"">" & Rec("Note").Value & "</td>"
							Ritorno &= "</tr>"
							Ritorno &= "<tr><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td></tr>"

							Dim idTP As String = ""
							If Rec("idTipoPagamento").Value Is DBNull.Value Then
								idTP = ""
							Else
								idTP = Rec("idTipoPagamento").Value
							End If
							If idTP = 1 Then
								totRata += Val(pag)
							Else
								totAltro += Val(pag)
							End If

							Dim metodo As String = "" & Rec("MetodoPagamento").Value
							Dim qm1 As Integer = 0
							Dim okm As Boolean = False
							For Each m As String In metodi
								If m = metodo Then
									totMetodi(qm1) += Val(pag)
									okm = True
									Exit For
								End If
								qm1 += 1
							Next
							If Not okm Then
								totMetodi.Add(Val(pag))
								metodi.Add(metodo)
							End If

							totale += Val(pag)

							Rec.MoveNext
						Loop
						Dim qm As Integer = 0
						For Each m As String In metodi
							Ritorno &= "<tr><td></td><td style=""text-align: left; font-weight: bold; padding-left: 3px;"">" & m & "</td><td style=""text-align: right; font-weight: bold;"">" & FormatCurrency(totMetodi(qm)) & "</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
							qm += 1
						Next
						Ritorno &= "<tr><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td></tr>"
						Ritorno &= "<tr><td></td><td style=""text-align: left; font-weight: bold; padding-left: 3px;"">Rata</td><td style=""text-align: right; font-weight: bold;"">" & FormatCurrency(totRata) & "</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
						Ritorno &= "<tr><td></td><td style=""text-align: left; font-weight: bold; padding-left: 3px;"">Altro</td><td style=""text-align: right; font-weight: bold;"">" & FormatCurrency(totAltro) & "</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
						Ritorno &= "<tr><td></td><td style=""text-align: left; font-weight: bold; padding-left: 3px;"">Totale</td><td style=""text-align: right; font-weight: bold;"">" & FormatCurrency(totale) & "</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
						Ritorno &= "</table>"

						'Ritorno &= "<hr /><div style=""text-algin: center; width: 100%;"">Stampato tramite InCalcio – www.incalcio.it – info@incalcio.it</div>"

						Rec.Close

						Dim gf As New GestioneFilesDirectory
						Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_lista_ricevute.txt")

						filetto = filetto.Replace("***TITOLO***", "Lista Ricevute")
						filetto = filetto.Replace("***DATI***", Ritorno)
						filetto = filetto.Replace("***NOME SQUADRA***", NomeSquadra)

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
						Dim pathLogo As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Societa_1.png"
						Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
						'Dim pathLogoConv As String = filePaths & "Appoggio\" & Esten & ".jpg"
						'Dim c As New CriptaFiles
						'c.DecryptFile(CryptPasswordString, pathLogo, pathLogoConv)

						Dim urlLogo As String = pathLogo ' mmPaths(2) & "Appoggio/" & Esten & ".jpg"
						filetto = filetto.Replace("***LOGO SOCIETA***", urlLogo)

						Dim nomeFileHtml As String = filePaths & "Appoggio\" & Esten & ".html"
						Dim nomeFilePDF As String = filePaths & "Appoggio\" & Esten & ".pdf"

						gf.CreaAggiornaFile(nomeFileHtml, filetto)

						Dim pp2 As New pdfGest
						Ritorno = pp2.ConverteHTMLInPDF(nomeFileHtml, nomeFilePDF, "",, True)
						If Ritorno = "*" Then
							Ritorno = "Appoggio/" & Esten & ".pdf"
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaInadempienti(Squadra As String) As String
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
				Dim Ok As Boolean = True

				Try
					Sql = "Select A.idQuota, Progressivo, Attiva, DescRata, DataScadenza, B.Descrizione, A.Importo From QuoteRate A Left Join Quote B On A.idQuota = B.idQuota " &
						"Where DataScadenza <> '' And DataScadenza Is Not Null And Convert(DateTime, DataScadenza ,121) <= getdate() And Attiva = 'S' " &
						"Order By Convert(DateTime, DataScadenza ,121)"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna rata quota rilevata"
						Else
							Dim quote As New List(Of String)

							Ritorno = ""
							Do Until Rec.Eof
								quote.Add(Rec("idQuota").Value & ";" & Rec("Progressivo").Value & ";" & Rec("DescRata").Value & ";" & Rec("Descrizione").Value & ";" & Rec("Importo").Value & ";" & Rec("DataScadenza").Value)

								Rec.MoveNext
							Loop
							Rec.Close

							For Each q As String In quote
								Dim qq() As String = q.Split(";")

								Sql = "Select A.* From Giocatori A " &
									"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
									"Left Join GiocatoriPagamenti C On A.idGiocatore = C.idGiocatore " &
									"Where B.idQuota = " & qq(0) & " And " &
									"(C.Progressivo Is Null Or (C.Progressivo In (" & qq(1) & ") And (C.NumeroRicevuta = 'Bozza' Or C.NumeroRicevuta Is Null)))"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2
								Else
									Do Until Rec2.Eof
										Ritorno &= Rec2("idGiocatore").Value & ";"
										Ritorno &= Rec2("Cognome").Value & ";"
										Ritorno &= Rec2("Nome").Value & ";"
										Ritorno &= qq(2) & ";"
										Ritorno &= qq(3) & ";"
										Ritorno &= qq(4) & ";"
										Ritorno &= qq(5) & ";"
										Ritorno &= Rec2("EMail").Value & ";"
										Ritorno &= "§"

										Rec2.MoveNext
									Loop
									Rec2.Close
								End If
							Next
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If
		End If
		If Ritorno = "" Then
			Ritorno = "Nessun giocatore inadempiente"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaRicevute(Squadra As String) As String
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
				Dim Ok As Boolean = True

				Try
					Sql = "Select A.idGiocatore, Progressivo, Pagamento, DataPagamento, B.Cognome, B.Nome, A.Validato, A.idTipoPagamento, " &
						"A.idRata, A.Note, A.idUtentePagatore, A.Commento, B.Maggiorenne, A.NumeroRicevuta, C.idMetodoPagamento, D.Cognome + ' ' + D.Nome As Nominativo, C.MetodoPagamento, E.idQuota From " &
						"GiocatoriPagamenti A " &
						"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
						"Left Join GiocatoriDettaglio E On A.idGiocatore = E.idGiocatore " &
						"Left Join MetodiPagamento C On A.MetodoPagamento = C.idMetodoPagamento " &
						"Left Join [Generale].[dbo].[Utenti] D On A.idUtenteRegistratore = D.idUtente " &
						"Where A.Eliminato = 'N' And B.Eliminato = 'N' " &
						"Order By NumeroRicevuta Desc" ' DataPagamento Desc, Progressivo Desc"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna ricevuta rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Dim nn As String = ("" & Rec("Note").Value).replace(";", "*PV*")
								Ritorno &= Rec("idGiocatore").Value & ";" & Rec("Progressivo").Value & ";" & Rec("Pagamento").Value & ";" &
									Rec("DataPagamento").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("Validato").Value & ";" &
									Rec("idTipoPagamento").Value & ";" & Rec("idRata").Value.replace(";", ":") & ";" & nn & ";" &
									Rec("idUtentePagatore").Value & ";" & Rec("Commento").Value & ";" & Rec("idQuota").Value & ";" & Rec("Maggiorenne").Value & ";" &
									Rec("NumeroRicevuta").Value & ";" & Rec("idMetodoPagamento").Value & ";" & Rec("Nominativo").Value & ";" &
									Rec("MetodoPagamento").Value & "§"

								Rec.MoveNext
							Loop
							Rec.Close
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
	Public Function RitornaQuote(Squadra As String) As String
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
				Dim Ok As Boolean = True

				Try
					Sql = "SELECT * FROM Quote Where Eliminato='N' Order By Descrizione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna quota rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idQuota").Value.ToString & ";"
								Ritorno &= Rec("Descrizione").Value.ToString & ";"
								Ritorno &= Rec("Importo").Value & ";"
								Ritorno &= Rec("Deducibilita").Value & ";"

								Sql = "Select * From QuoteRate Where idQuota=" & Rec("idQuota").Value & " And Eliminato='N' Order By Progressivo"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2

									Ritorno &= "N;;;;"
									Ritorno &= "N;;;;"
									Ritorno &= "N;;;;"
									Ritorno &= "N;;;;"
									Ritorno &= "N;;;;"

									Ok = False
									Exit Do
								Else
									Dim q As Integer = 0

									Do Until Rec2.Eof
										Ritorno &= Rec2("Attiva").Value & ";"
										Ritorno &= Rec2("DescRata").Value & ";"
										Ritorno &= Rec2("DataScadenza").Value & ";"
										Ritorno &= Rec2("Importo").Value & ";"
										q += 1

										Rec2.MoVeNext
									Loop
									Rec2.Close()

									For i As Integer = q To 5
										Ritorno &= "N;;;;"
									Next
								End If

								Sql = "Select * From GiocatoriPagamenti Where idQuota = " & Rec("idQuota").Value & " And Eliminato='N'"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ritorno = Rec2

									Ok = False
									Exit Do
								Else
									If Rec2.Eof() Then
										Ritorno &= "S;"
									Else
										Ritorno &= "N;"
									End If
								End If

								Ritorno &= "§"

								Rec.MoveNext()

								If Not Ok Then
									Exit Do
								End If
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
	Public Function EliminaQuota(Squadra As String, ByVal idQuota As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Giocata As String = ""

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
						Sql = "Update Quote Set Eliminato='S' " &
								"Where idQuota=" & idQuota
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
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaQuota(Squadra As String, ByVal idQuota As String, Descrizione As String, Importo As String,
								   AttivaR1 As String, DescRataR1 As String, DataScadenzaR1 As String, ImportoR1 As String,
								   AttivaR2 As String, DescRataR2 As String, DataScadenzaR2 As String, ImportoR2 As String,
								   AttivaR3 As String, DescRataR3 As String, DataScadenzaR3 As String, ImportoR3 As String,
								   AttivaR4 As String, DescRataR4 As String, DataScadenzaR4 As String, ImportoR4 As String,
								   AttivaR5 As String, DescRataR5 As String, DataScadenzaR5 As String, ImportoR5 As String,
								   Deducibilita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Giocata As String = ""

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
						Sql = "Update Quote Set " &
							"Descrizione='" & Descrizione.Replace("'", "''") & "', " &
							"Importo=" & Importo & ", " &
							"Deducibilita='" & Deducibilita & "' " &
							"Where idQuota=" & idQuota
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						Else
							Sql = "Update QuoteRate Set " &
								"Attiva='" & AttivaR1 & "', " &
								"DescRata='" & DescRataR1.Replace("'", "''") & "', " &
								"DataScadenza='" & DataScadenzaR1 & "', " &
								"Importo=" & ImportoR1 & " " &
								"Where idQuota=" & idQuota & " And Progressivo=1"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							Else
								Sql = "Update QuoteRate Set " &
									"Attiva='" & AttivaR2 & "', " &
									"DescRata='" & DescRataR2.Replace("'", "''") & "', " &
									"DataScadenza='" & DataScadenzaR2 & "', " &
									"Importo=" & ImportoR2 & " " &
									"Where idQuota=" & idQuota & " And Progressivo=2"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								Else
									Sql = "Update QuoteRate Set " &
										"Attiva='" & AttivaR3 & "', " &
										"DescRata='" & DescRataR3.Replace("'", "''") & "', " &
										"DataScadenza='" & DataScadenzaR3 & "', " &
										"Importo=" & ImportoR3 & " " &
										"Where idQuota=" & idQuota & " And Progressivo=3"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									Else
										Sql = "Update QuoteRate Set " &
											"Attiva='" & AttivaR4 & "', " &
											"DescRata='" & DescRataR4.Replace("'", "''") & "', " &
											"DataScadenza='" & DataScadenzaR4 & "', " &
											"Importo=" & ImportoR4 & " " &
											"Where idQuota=" & idQuota & " And Progressivo=4"
										Ritorno = EsegueSql(Conn, Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										Else
											Sql = "Update QuoteRate Set " &
												"Attiva='" & AttivaR5 & "', " &
												"DescRata='" & DescRataR5.Replace("'", "''") & "', " &
												"DataScadenza='" & DataScadenzaR5 & "', " &
												"Importo=" & ImportoR5 & " " &
												"Where idQuota=" & idQuota & " And Progressivo=5"
											Ritorno = EsegueSql(Conn, Sql, Connessione)
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
											End If
										End If
									End If
								End If
							End If
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
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InserisceQuota(Squadra As String, Descrizione As String, Importo As String,
								   AttivaR1 As String, DescRataR1 As String, DataScadenzaR1 As String, ImportoR1 As String,
								   AttivaR2 As String, DescRataR2 As String, DataScadenzaR2 As String, ImportoR2 As String,
								   AttivaR3 As String, DescRataR3 As String, DataScadenzaR3 As String, ImportoR3 As String,
								   AttivaR4 As String, DescRataR4 As String, DataScadenzaR4 As String, ImportoR4 As String,
								   AttivaR5 As String, DescRataR5 As String, DataScadenzaR5 As String, ImportoR5 As String,
								   Deducibilita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Giocata As String = ""

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

				Dim idQuota As Integer = -1

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "SELECT Max(idQuota)+1 FROM Quote"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec(0).Value Is DBNull.Value Then
								idQuota = 1
							Else
								idQuota = Rec(0).Value
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						Try
							Sql = "Insert Into Quote Values (" & idQuota & ", '" & Descrizione.Replace("'", "''") & "', " & Importo & ", 'N', '" & Deducibilita & "')"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							Else
								Sql = "Insert Into QuoteRate Values (" &
									" " & idQuota & ", " &
									"1, " &
									"'" & AttivaR1 & "', " &
									"'" & DescRataR1.Replace("'", "''") & "', " &
									"'" & DataScadenzaR1 & "', " &
									" " & ImportoR1 & ", " &
									"'N' " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								Else
									Sql = "Insert Into QuoteRate Values (" &
										" " & idQuota & ", " &
										"2, " &
										"'" & AttivaR2 & "', " &
										"'" & DescRataR2.Replace("'", "''") & "', " &
										"'" & DataScadenzaR2 & "', " &
										" " & ImportoR2 & ", " &
										"'N' " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									Else
										Sql = "Insert Into QuoteRate Values (" &
											" " & idQuota & ", " &
											"3, " &
											"'" & AttivaR3 & "', " &
											"'" & DescRataR3.Replace("'", "''") & "', " &
											"'" & DataScadenzaR3 & "', " &
											" " & ImportoR3 & ", " &
											"'N' " &
											")"
										Ritorno = EsegueSql(Conn, Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										Else
											Sql = "Insert Into QuoteRate Values (" &
												" " & idQuota & ", " &
												"4, " &
												"'" & AttivaR4 & "', " &
												"'" & DescRataR4.Replace("'", "''") & "', " &
												"'" & DataScadenzaR4 & "', " &
												" " & ImportoR4 & ", " &
												"'N' " &
												")"
											Ritorno = EsegueSql(Conn, Sql, Connessione)
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
											Else
												Sql = "Insert Into QuoteRate Values (" &
													" " & idQuota & ", " &
													"5, " &
													"'" & AttivaR5 & "', " &
													"'" & DescRataR5.Replace("'", "''") & "', " &
													"'" & DataScadenzaR5 & "', " &
													" " & ImportoR5 & ", " &
													"'N' " &
													")"
												Ritorno = EsegueSql(Conn, Sql, Connessione)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
												End If
											End If
										End If
									End If
								End If
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
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function
End Class