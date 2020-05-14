Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_uteloc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsUtentiLocali
	Inherits System.Web.Services.WebService


	<WebMethod()>
	Public Function RitornaUtentePerLoginNuovo(Utente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim Squadra As String = ""
		Dim UtenteDaSalvare As String = ""

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
					Sql = "SELECT Utenti.idAnno, idUtente, Utente, Cognome, Nome, " &
						"Password, EMail, idCategoria, idTipologia, Utenti.idSquadra, Descrizione As Squadra " &
						"FROM Utenti Left Join Squadre On Utenti.idSquadra = Squadre.idSquadra " &
						"Where Upper(Utente)='" & Utente.ToUpper.Replace("'", "''") & "'"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							'If Password <> DecriptaStringa(Rec("Password").Value.ToString) Then
							'	Ritorno = StringaErrore & " Password non valida"
							'Else
							Ritorno = ""
								Do Until Rec.Eof
									Ritorno &= Rec("idAnno").Value & ";" &
										Rec("idUtente").Value & ";" &
										Rec("Utente").Value & ";" &
										Rec("Cognome").Value & ";" &
										Rec("Nome").Value & ";" &
										DecriptaStringa(Rec("Password").Value) & ";" &
										Rec("EMail").Value & ";" &
										Rec("idCategoria").Value & ";" &
										Rec("idTipologia").Value & ";" &
										Rec("idSquadra").Value & ";" &
										Rec("Squadra").Value & ";" &
										"§"

									Squadra = Rec("Squadra").Value
									UtenteDaSalvare = Ritorno

									Rec.MoveNext()
								Loop
							'End If
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()

				If Not Ritorno.Contains(StringaErrore) Then
					Dim Connessione2 As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra.Replace(" ", "_"))

					If Connessione2 = "" Then
						Ritorno = ErroreConnessioneNonValida
					Else
						Dim Conn2 As Object = ApreDB(Connessione2)
						Dim Ritorno2 As String = ""

						If TypeOf (Conn) Is String Then
							Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
						Else
							Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
							Dim Sql2 As String = ""
							Dim Campi() As String = UtenteDaSalvare.Split(";")

							Sql2 = "Select * From Utenti Where Upper(Utente) = '" & Utente.ToUpper.Replace("'", "''") & "' And idAnno=" & Campi(0)
							Rec2 = LeggeQuery(Conn2, Sql2, Connessione2)
							If TypeOf (Rec) Is String Then
								' Ritorno = Rec2
							Else
								If Rec2.Eof Then
									' Aggiungo l'utente rilevato nel db generale e non in quello di lavoro
									Sql2 = "Insert Into Utenti Values (" &
										" " & Campi(0) & ", " &
										" " & Campi(1) & ", " &
										"'" & Campi(2).Replace("'", "''") & "', " &
										"'" & Campi(3).Replace("'", "''") & "', " &
										"'" & Campi(4).Replace("'", "''") & "', " &
										"'" & CriptaStringa(Campi(5)).Replace("'", "''") & "', " &
										"'" & Campi(6).Replace("'", "''") & "', " &
										" " & Campi(7) & ", " &
										" " & Campi(8) & " " &
										")"
									Ritorno2 = EsegueSql(Conn2, Sql2, Connessione2)

									If Not Ritorno2.Contains(StringaErrore) Then

									End If
								End If
								Rec2.Close()

							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaUtentePerLogin(Squadra As String, ByVal idAnno As String, Utente As String, Password As String) As String
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
					' Sql = "SELECT * FROM Utenti Where idAnno=" & idAnno & " And Utente='" & Utente.Replace("'", "''") & "'"
					Sql = "SELECT Utenti.idAnno, idUtente, Utente, Cognome, Nome, " &
						"Password, EMail, Categorie.idCategoria As idCat1, idTipologia, Categorie.Descrizione As Descr1 " &
						"FROM (Utenti " &
						"Left Join Categorie On Utenti.idCategoria=Categorie.idCategoria And Utenti.idAnno=Categorie.idAnno) " &
						"Where Utente='" & Utente.Replace("'", "''") & "' And Utenti.idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							If Password <> DecriptaStringa(Rec("Password").Value.ToString) Then
								Ritorno = StringaErrore & " Password non valida"
							Else
								Ritorno = ""
								Do Until Rec.Eof
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
					' Sql = "SELECT * FROM Utenti Where idAnno=" & idAnno & " And idUtente=" & idUtente
					Sql = "SELECT Utenti.*, Categorie.Descrizione " &
						"From Utenti LEFT Join Categorie On (Utenti.idCategoria = Categorie.idCategoria) And (Utenti.idAnno = Categorie.idAnno) " &
						"Where idUtente = " & idUtente
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
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
	Public Function RitornaListaUtenti(Squadra As String, idAnno As String) As String
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
					Sql = "SELECT Utenti.idAnno, Utenti.idUtente, Utenti.Utente, Utenti.Cognome, Utenti.Nome, Utenti.EMail, Categorie.Descrizione As Categoria, " &
						"Utenti.idTipologia, Utenti.Password, Categorie.idCategoria " &
						"FROM (Utenti LEFT JOIN Categorie ON Utenti.idCategoria = Categorie.idCategoria And Utenti.idAnno = Categorie.idAnno) " &
						"Where Utenti.idAnno=" & idAnno & " Order By 2,1;"
					' "Where Utenti.idAnno=" & idAnno & " Order By 2,1;"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")

								Sql = " Select * From AnnoAttualeUtenti Where idUtente=" & Rec("idUtente").Value
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								Dim AnnoUtente As Integer = Rec2("idAnno").Value
								Rec2.Close

								Sql = " Select * From Anni Where idAnno=" & AnnoUtente.ToString
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								Dim NomeSquadra As String = Rec2("NomeSquadra").Value
								Rec2.Close

								Ritorno &= "0;" &
									Rec("idUtente").Value & ";" &
									Rec("Utente").Value & ";" &
									Rec("Cognome").Value & ";" &
									Rec("Nome").Value & ";" &
									Rec("EMail").Value & ";" &
									NomeSquadra & ";" &
									Rec("idTipologia").Value & ";" &
									DecriptaStringa(Rec("Password").Value) & ";" &
									Rec("idCategoria").Value & ";" &
									Rec("Categoria").Value & ";" &
									"§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()

						'Ritorno &= "£"

						'Sql = "Select * From Categorie Where idAnno=" & idAnno & " And Eliminato = 'N' Order By Ordinamento"
						'Rec = LeggeQuery(Conn, Sql, Connessione)
						'If TypeOf (Rec) Is String Then
						'    Ritorno = Rec
						'Else
						'    If Rec.Eof Then
						'        Ritorno = StringaErrore & " Nessuna categoria rilevata"
						'    Else
						'        Do Until Rec.Eof
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
	Public Function SalvaUtente(Squadra As String, ByVal idAnno As String, Utente As String, Cognome As String, Nome As String, EMail As String,
								Password As String, idCategoria As String, idTipologia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim idUtente As String = ""

				Try
					Sql = "SELECT * FROM Utenti Where Upper(Utente)='" & Utente.Trim.ToUpper & "' And idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Sql = "SELECT Max(idUtente)+1 FROM Utenti Where idAnno=" & idAnno
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = StringaErrore & " Nessun utente rilevato"
								Else
									If Rec(0).Value Is DBNull.Value Then
										idUtente = "1"
									Else
										idUtente = Rec(0).Value.ToString
									End If
								End If
								Rec.Close()
							End If

							If idUtente <> "" Then
								Dim idSquadra As Integer

								Sql = "Select idSquadra From Squadre Where Descrizione='" & Squadra.Replace("_", " ").Replace("'", "''") & "'"
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec.Eof Then
										Ritorno = StringaErrore & " Nessuna squadra rilevata"
									Else
										idSquadra = Rec(0).Value
										Rec.Close()

										Sql = "Insert Into Utenti Values (" &
											" " & idAnno & ", " &
											" " & idUtente & ", " &
											"'" & Utente.Replace("'", "''") & "', " &
											"'" & Cognome.Replace("'", "''") & "', " &
											"'" & Nome.Replace("'", "''") & "', " &
											"'" & CriptaStringa(Password).Replace("'", "''") & "', " &
											"'" & EMail.Replace("'", "''") & "', " &
											" " & idCategoria & ", " &
											" " & idTipologia & ", " &
											" " & idSquadra & " " &
											")"
										Ritorno = EsegueSql(Conn, Sql, Connessione)
									End If
								End If

								If Ritorno = "*" Then
									Ritorno = idUtente

									Dim Connessione2 As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
									Dim Ritorno2 As String = ""

									If Connessione2 = "" Then
										Ritorno2 = ErroreConnessioneNonValida
									Else
										Dim Conn2 As Object = ApreDB(Connessione2)

										If TypeOf (Conn2) Is String Then
											Ritorno2 = ErroreConnessioneDBNonValida & ":" & Conn2
										Else
											Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
											Dim Sql2 As String = ""

											Sql2 = "Insert Into Utenti Values (" &
												" " & idAnno & ", " &
												" " & idUtente & ", " &
												"'" & Utente.Replace("'", "''") & "', " &
												"'" & Cognome.Replace("'", "''") & "', " &
												"'" & Nome.Replace("'", "''") & "', " &
												"'" & CriptaStringa(Password).Replace("'", "''") & "', " &
												"'" & EMail.Replace("'", "''") & "', " &
												" " & idCategoria & ", " &
												" " & idTipologia & " " &
												")"
											Ritorno2 = EsegueSql(Conn2, Sql2, Connessione2)

											If Ritorno2 <> "*" Then
												Ritorno = Ritorno2
											End If
										End If
									End If
								End If
							Else
								Ritorno = StringaErrore & " Problemi nel rilevamento dell'ID Utente"
							End If
						Else
							Ritorno = StringaErrore & " Utente già esistente per l'anno in corso"
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
	Public Function EliminaUtente(Squadra As String, ByVal idAnno As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Sql = "Delete From Utenti Where idUtente=" & idUtente & " And idAnno=" & idAnno
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Ritorno = "*" Then
					Conn.Close()

					Connessione = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

					If Connessione = "" Then
						Ritorno = ErroreConnessioneNonValida
					Else
						Conn = ApreDB(Connessione)

						If TypeOf (Conn) Is String Then
							Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
						Else
							Sql = "Delete From Utenti Where idUtente=" & idUtente & " And idAnno=" & idAnno
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						End If
					End If
				End If
					End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaUtente(Squadra As String, ByVal idAnno As String, Utente As String, Cognome As String, Nome As String, EMail As String,
								Password As String, idCategoria As String, idTipologia As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

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

				' Sql = "Delete From Utenti Where idAnno=" & idAnno & " And idUtente=" & idUtente
				Sql = "Delete From Utenti Where idUtente=" & idUtente & " And idAnno=" & idAnno
				Ritorno = EsegueSql(Conn, Sql, Connessione)
				If Ritorno.Contains(StringaErrore) Then
					Ok = False
				End If

				If Ok Then
					Dim idSquadra As Integer

					Sql = "Select idSquadra From Squadre Where Descrizione='" & Squadra.Replace("_", " ").Replace("'", "''") & "'"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna squadra rilevata"
						Else
							idSquadra = Rec(0).Value
							Rec.Close()

							Try
								Sql = "Insert Into Utenti Values (" &
									"" & idAnno & ", " &
									"" & idUtente & ", " &
									"'" & Utente.Replace("'", "''") & "', " &
									"'" & Cognome.Replace("'", "''") & "', " &
									"'" & Nome.Replace("'", "''") & "', " &
									"'" & CriptaStringa(Password).Replace("'", "''") & "', " &
									"'" & EMail.Replace("'", "''") & "', " &
									" " & idCategoria & ", " &
									"" & idTipologia & ", " &
									"" & idSquadra & ")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

								If Ok Then
									Conn.Close

									Connessione = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

									If Connessione = "" Then
										Ritorno = ErroreConnessioneNonValida
									Else
										Conn = ApreDB(Connessione)

										If TypeOf (Conn) Is String Then
											Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
										Else
											Try
												Sql = "Delete From Utenti Where idUtente=" & idUtente & " And idAnno=" & idAnno
												Ritorno = EsegueSql(Conn, Sql, Connessione)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
												End If

												If Ok Then
													Sql = "Insert Into Utenti Values (" &
														"" & idAnno & ", " &
														"" & idUtente & ", " &
														"'" & Utente.Replace("'", "''") & "', " &
														"'" & Cognome.Replace("'", "''") & "', " &
														"'" & Nome.Replace("'", "''") & "', " &
														"'" & CriptaStringa(Password).Replace("'", "''") & "', " &
														"'" & EMail.Replace("'", "''") & "', " &
														" " & idCategoria & ", " &
														" " & idTipologia & " )"
													Ritorno = EsegueSql(Conn, Sql, Connessione)

													If Ritorno <> "*" Then
														Ok = False
													End If

													If Ok Then
														Sql = "Delete From AnnoAttualeUtenti Where idUtente=" & idUtente
														Ritorno = EsegueSql(Conn, Sql, Connessione)
														If Ritorno.Contains(StringaErrore) Then
															Ok = False
														End If

														If Ok Then
															Sql = "Insert Into AnnoAttualeUtenti Values (" & idUtente & ", " & idAnno & ")"
															Ritorno = EsegueSql(Conn, Sql, Connessione)
															If Ritorno.Contains(StringaErrore) Then
																Ok = False
															End If

														End If
													End If
												End If
											Catch ex As Exception
												Ritorno = StringaErrore & " " & ex.Message
											End Try
										End If
									End If
								End If

								If Ritorno = "*" Then Ritorno = idUtente
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
							End Try
						End If
					End If

					Conn.Close()
				End If
			End If
		End If

		Return Ritorno
	End Function
End Class