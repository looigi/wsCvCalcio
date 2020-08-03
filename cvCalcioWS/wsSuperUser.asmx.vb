Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Runtime.CompilerServices

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://cvcalcio_super_user.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsSuperUser
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function CreaDB(Squadra As String, DataScadenza As String, MailAdmin As String, NomeAdmin As String, CognomeAdmin As String, Anno As String, idTipologia As String, idLicenza As String) As String
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
							End If

							If Ok Then
								Dim Societa As String = Format(idSocieta, "00000")
								nomeDb = "0001_" & Societa

								Sql = "Create Database [" & nomeDb & "]"
								Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

								If Ok Then
									Sql = "Begin transaction"
									Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

									Sql = "Insert Into Squadre Values (" & idSocieta & ", '" & Squadra.Replace("'", "''") & "', '" & DataScadenza & "', " & idTipologia & ", " & idLicenza & ", 'N')"
									Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If

									If Ok Then
										Sql = "Insert Into Utenti Values (" &
											"1, " &
											" " & idUtente & ", " &
											"'" & MailAdmin.Replace("'", "''") & "', " &
											"'" & CognomeAdmin.Replace("'", "''") & "', " &
											"'" & NomeAdmin.Replace("'", "''") & "', " &
											"'" & CriptaStringa("Password123!") & "', " &
											"'" & MailAdmin.Replace("'", "''") & "', " &
											"-1, " &
											"1, " &
											" " & idSocieta & ", " &
											"1, " &
											"'', " &
											"'N' " &
											")"
										Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If

										If Ok Then
											For i As Integer = 0 To q - 1
												Try
													Sql = "Select * Into [" & nomeDb & "].[dbo].[" & Tabelle(i) & "] From " & Tabelle(i)
													Ritorno = EsegueSql(ConnDbVuoto, Sql, ConnessioneDBVuoto)
													If Ritorno.Contains(StringaErrore) Then
														Ok = False
													End If
												Catch ex As Exception
													Ritorno = StringaErrore & " " & ex.Message
													Ok = False
												End Try
											Next

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
													"null " &
													")"
												Ritorno = EsegueSql(ConnDbVuoto, Sql, ConnessioneDBVuoto)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
												Else
													Sql = "Insert Into SquadraAnni Values (" &
														" " & idSocieta & ", " &
														"1, " &
														"'" & Anno & "' " &
														")"
													Ritorno = EsegueSql(ConnGen, Sql, ConnessioneGenerale)
													If Ritorno.Contains(StringaErrore) Then
														Ok = False
													Else
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

															Body &= "E' stata creata la nuova società '" & Squadra & "'. " & vbCrLf & vbCrLf
															Body &= "Amministratore: " & CognomeAdmin & " " & NomeAdmin & vbCrLf
															Body &= "Scadenza licenza: " & DataScadenza & vbCrLf
															Body &= "Anno: " & Anno & vbCrLf
															Body &= "Tipologia: " & Tipologia & vbCrLf
															Body &= "Licenza: " & idLicenza & vbCrLf

															Dim ChiScrive As String = "notifiche@incalcio.cloud"

															Ritorno = m.SendEmail(Oggetto, Body, ChiScrive, MailAdmin)
															If Not Ritorno.Contains(StringaErrore) Then
																Ritorno = Societa
															End If
														Catch ex As Exception
															Ritorno = StringaErrore & " " & ex.Message
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
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(ConnGen, Sql, ConnessioneGenerale)

					Sql = "Drop Database [" & nomeDb & "]"
					Ritorno2 = EsegueSql(ConnDbVuoto, Sql, ConnessioneDBVuoto)
				End If
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
									maxAnno = Rec2(2).Value
								End If
								Rec2.Close
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
End Class