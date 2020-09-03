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
									"Where B.idQuota = " & qq(0) & " And (C.Progressivo Not In (" & qq(1) & ") Or C.Progressivo Is Null)"
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
					Sql = "Select A.idGiocatore, Progressivo, Pagamento, DataPagamento, B.Cognome, B.Nome, A.Validato, A.idTipoPagamento, A.idRata, A.Note, A.idUtentePagatore, A.Commento, B.Maggiorenne From " &
						"GiocatoriPagamenti A " &
						"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
						"Where B.Eliminato = 'N' " &
						"Order By DataPagamento Desc"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna ricevuta rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idGiocatore").Value & ";" & Rec("Progressivo").Value & ";" & Rec("Pagamento").Value & ";" &
									Rec("DataPagamento").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("Validato").Value & ";" &
									Rec("idTipoPagamento").Value & ";" & Rec("idRata").Value & ";" & Rec("Note").Value.replace(";", "*PV*") & ";" &
									Rec("idUtentePagatore").Value & ";" & Rec("Commento").Value & ";" & Rec("Maggiorenne").Value & ";" &
									"§"

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
								   AttivaR5 As String, DescRataR5 As String, DataScadenzaR5 As String, ImportoR5 As String) As String
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
							Sql = "Insert Into Quote Values (" & idQuota & ", '" & Descrizione.Replace("'", "''") & "', " & Importo & ", 'N')"
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