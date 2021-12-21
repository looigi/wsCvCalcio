Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_cat.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsCategorie
    Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaNuovoID(Squadra As String, ByVal idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idCategoria As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				'Dim idUtente As String = ""

				Sql = "SELECT Max(idCategoria)+1 FROM Categorie Where idAnno=" & idAnno
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec(0).Value Is DBNull.Value Then
						idCategoria = "1"
					Else
						idCategoria = Rec(0).Value.ToString
					End If
				End If
				Rec.Close()
			End If
		End If

		Return idCategoria
	End Function

	<WebMethod()>
	Public Function SalvaCategorieUtente(Squadra As String, IDutente As Integer, Categorie As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Ok As Boolean = True
		Dim Sql As String = ""

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Dim TipoUtente As String = ""
				Dim NomeTabella As String = ""

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & IDutente
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Nessun utente rilevato"
							Else
								TipoUtente = Rec("idTipologia").Value
							End If
							Rec.Close()
						End If

						Select Case TipoUtente
							Case "5", "7"
								NomeTabella = "AllenatoriCategorie"
							Case "4", "8"
								NomeTabella = "DirigentiCategorie"
							Case Else
								NomeTabella = "UtentiCategorie"
						End Select

						Sql = "Delete From " & NomeTabella & " Where Idutente = " & IDutente
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						If Categorie.Length > 0 Then
							Dim Cate() As String = Categorie.Split(",")
							Dim Progressivo As Integer = 0

							For Each p As String In Cate
								If p <> "" Then
									Progressivo += 1

									Try
										Sql = "Insert Into " & NomeTabella & " Values (" &
											" " & IDutente & ", " &
											" " & Progressivo & ", " &
											" " & p & " " &
											")"
										Ritorno = EsegueSql(Conn, Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
											Exit For
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & ex.Message
										Ok = False
										Exit For
									End Try
								End If
							Next

						End If
					End If
				End If
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

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaTutteCategorieUtente(Squadra As String, idAnno As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim ConnessioneU As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim ConnU As Object = ApreDB(ConnessioneU)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				If TypeOf (ConnU) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
				Else
					Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
					Dim Sql As String = ""
					Dim TipoUtente As String = ""

					Try
						Sql = "Select * From Utenti Where idUtente=" & idUtente
						Rec = LeggeQuery(ConnU, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Nessun utente rilevato"
							Else
								TipoUtente = Rec("idTipologia").Value
							End If
							Rec.Close()
						End If
					Catch ex As Exception

					End Try

					If Ritorno = "" Then
						Try
							Select Case TipoUtente
								Case "2"
									Sql = "SELECT * From UtentiCategorie Where idUtente=" & idUtente
								Case "5", "7"
									Sql = "SELECT * From AllenatoriCategorie Where idUtente=" & idUtente
								Case "4", "8"
									Sql = "SELECT * From DirigentiCategorie Where idUtente=" & idUtente
								Case Else
									Sql = "SELECT * From Categorie Where idAnno=" & idAnno
							End Select
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									' Ritorno = StringaErrore & " Nessun permesso ritornato"
								Else
									Ritorno = ""
									Do Until Rec.Eof
										Ritorno &= Rec("idCategoria").Value.ToString & "§"

										Rec.MoveNext()
									Loop
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							'				Ritorno = StringaErrore & " " & ex.Message
						End Try
					End If

					ConnU.Close()
					Conn.Close()
				End If
			End If

		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaCategorie(Squadra As String, ByVal idAnno As String, idUtente As String) As String
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
					Sql = "Select * From [Generale].[dbo].Utenti Where idAnno=" & idAnno & " And idUtente=" & idUtente
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							Dim idTipologia As String = Rec("idTipologia").Value
							Dim idCategoria As String = Rec("idCategoria").Value

							'{ id 1, Descrizione: 'Amministratore' }, 
							'{ id 2, Descrizione: 'Utente' }, 
							'{ id 3, Descrizione: 'Genitore'}, 
							'{ id 4, Descrizione: 'Dirigente Genitore'}, 
							'{ id 5, Descrizione: 'Allenatore Genitore'},
							'{ id 6, Descrizione: 'Giocatore'},
							'{ id 7, Descrizione: 'Allenatore'},
							'{ id 8, Descrizione 'Dirigente'},

							Select Case idTipologia
								Case "0", "1"
									Sql = "SELECT idCategoria, Descrizione, AnticipoConvocazione, RisultatoATempi, " &
										"GiornoAllenamento1, OraInizio1, OraFine1, " &
										"GiornoAllenamento2, OraInizio2, OraFine2, " &
										"GiornoAllenamento3, OraInizio3, OraFine3, " &
										"GiornoAllenamento4, OraInizio4, OraFine4, " &
										"AnnoCategoria, ShootOut, Tempi, TempoGioco " &
										"FROM Categorie Where idAnno=" & idAnno & " And Eliminato='N' Order By Descrizione"
								Case "5", "7"
									Sql = "Select A.idCategoria, B.Descrizione, AnticipoConvocazione, B.RisultatoATempi, " &
										"B.GiornoAllenamento1, B.OraInizio1, B.OraFine1, " &
										"B.GiornoAllenamento2, B.OraInizio2, B.OraFine2, " &
										"B.GiornoAllenamento3, B.OraInizio3, B.OraFine3, " &
										"B.GiornoAllenamento4, B.OraInizio4, B.OraFine4, " &
										"AnnoCategoria, ShootOut, Tempi, TempoGioco " &
										"From AllenatoriCategorie A " &
										"Left Join Categorie B On A.idCategoria = B.idCategoria " &
										"Where B.idAnno = " & idAnno & " And A.idUtente = " & idUtente & " And Eliminato='N' Order By Descrizione"
								Case "4", "8"
									Sql = "Select A.idCategoria, B.Descrizione, AnticipoConvocazione, B.RisultatoATempi, " &
										"B.GiornoAllenamento1, B.OraInizio1, B.OraFine1, " &
										"B.GiornoAllenamento2, B.OraInizio2, B.OraFine2, " &
										"B.GiornoAllenamento3, B.OraInizio3, B.OraFine3, " &
										"B.GiornoAllenamento4, B.OraInizio4, B.OraFine4, " &
										"AnnoCategoria, ShootOut, Tempi, TempoGioco " &
										"From DirigentiCategorie A " &
										"Left Join Categorie B On A.idCategoria = B.idCategoria " &
										"Where B.idAnno = " & idAnno & " And A.idUtente = " & idUtente & " And Eliminato='N' Order By Descrizione"
								Case Else
									Sql = "Select A.idCategoria, B.Descrizione, AnticipoConvocazione, B.RisultatoATempi, " &
										"B.GiornoAllenamento1, B.OraInizio1, B.OraFine1, " &
										"B.GiornoAllenamento2, B.OraInizio2, B.OraFine2, " &
										"B.GiornoAllenamento3, B.OraInizio3, B.OraFine3, " &
										"B.GiornoAllenamento4, B.OraInizio4, B.OraFine4, " &
										"AnnoCategoria, ShootOut, Tempi, TempoGioco " &
										"From UtentiCategorie A " &
										"Left Join Categorie B On A.idCategoria = B.idCategoria " &
										"Where B.idAnno = " & idAnno & " And A.idUtente = " & idUtente & " And Eliminato='N' Order By Descrizione"
							End Select

							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = "" ' StringaErrore & " Nessuna categoria rilevata"
								Else
									Ritorno = ""
									If idTipologia = "1" Or idTipologia = "0" Then
										Ritorno &= "-1;Tutte le categorie;0§"
									End If
									Do Until Rec.Eof
										Ritorno &= Rec("idCategoria").Value.ToString & ";" & Rec("Descrizione").Value.ToString & ";" &
											Rec("AnticipoConvocazione").Value & ";" & Rec("RisultatoATempi").Value & ";" &
											Rec("GiornoAllenamento1").Value & ";" & Rec("OraInizio1").Value & ";" & Rec("OraFine1").Value & ";" &
											Rec("GiornoAllenamento2").Value & ";" & Rec("OraInizio2").Value & ";" & Rec("OraFine2").Value & ";" &
											Rec("GiornoAllenamento3").Value & ";" & Rec("OraInizio3").Value & ";" & Rec("OraFine3").Value & ";" &
											Rec("GiornoAllenamento4").Value & ";" & Rec("OraInizio4").Value & ";" & Rec("OraFine4").Value & ";" &
											Rec("AnnoCategoria").Value & ";" & Rec("ShootOut").Value & ";" & Rec("Tempi").Value & ";" & Rec("TempoGioco").Value &
											"§"

										Rec.MoveNext()
									Loop
								End If
								Rec.Close()
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
	Public Function RitornaCategoriePerAnno(Squadra As String, ByVal idAnno As String) As String
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
					Sql = "SELECT * FROM Categorie Where idAnno=" & idAnno & " And Eliminato='N' Order By Descrizione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = "" ' StringaErrore & " Nessuna categoria rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idCategoria").Value.ToString & ";" & Rec("Descrizione").Value.ToString & ";" & Rec("AnticipoConvocazione").Value & ";" & Rec("RisultatoATempi").Value & ";" & Rec("AnnoCategoria").Value & ";" & Rec("TempoGioco").Value & "§"

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
	Public Function SalvaCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String, Categoria As String, AnticipoConvocazione As String, RisultatoATempi As String,
								   GiornoAllenamento1 As String, OraInizio1 As String, OraFine1 As String, GiornoAllenamento2 As String, OraInizio2 As String, OraFine2 As String,
								   GiornoAllenamento3 As String, OraInizio3 As String, OraFine3 As String, GiornoAllenamento4 As String, OraInizio4 As String, OraFine4 As String,
								   AnnoCategoria As String, ShootOut As String, Tempi As String, TempoGioco As String) As String
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
					If idCategoria = -1 Then
						Dim Rec As Object = Server.CreateObject("ADODB.Recordset")

						Try
							Sql = "Select Max(idCategoria)+1 From Categorie Where idAnno=" & idAnno
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec(0).Value Is DBNull.Value Then
									idCategoria = 1
								Else
									idCategoria = Rec(0).Value
								End If
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					Else
						Try
							Sql = "Delete From Categorie Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
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
						Try
							Sql = "Insert Into Categorie Values (" &
								" " & idAnno & ", " &
								" " & idCategoria & ", " &
								"'" & Categoria.Replace("'", "''") & "', " &
								"'N', " &
								"1," &
								" " & AnticipoConvocazione & ", " &
								"'" & RisultatoATempi & "', " &
								" " & GiornoAllenamento1 & ", " &
								"'" & OraInizio1 & "', " &
								"'" & OraFine1 & "', " &
								" " & GiornoAllenamento2 & ", " &
								"'" & OraInizio2 & "', " &
								"'" & OraFine2 & "', " &
								" " & GiornoAllenamento3 & ", " &
								"'" & OraInizio3 & "', " &
								"'" & OraFine3 & "', " &
								" " & GiornoAllenamento4 & ", " &
								"'" & OraInizio4 & "', " &
								"'" & OraFine4 & "', " &
								"'" & AnnoCategoria & "', " &
								"'" & ShootOut & "', " &
								" " & Tempi & ", " &
								" " & TempoGioco & " " &
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
	Public Function EliminaCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String) As String
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

				Try
					Sql = "Update Categorie Set Eliminato='S' Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
					Ritorno = EsegueSql(Conn, Sql, Connessione)
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

End Class