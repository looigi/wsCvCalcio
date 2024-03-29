﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_ute.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsUtenti
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function CriptaDecriptaStringaPerProva(Stringa As String) As String
		Dim s As String = CriptaStringa(Stringa)
		Dim r As String = DecriptaStringa(s)

		Return s
	End Function

	Public Function RitornaUtentePerLogin(Squadra As String, ByVal idAnno As String, Utente As String, Password As String) As String
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
					' Sql = "SELECT * FROM UtentiMobile Where idAnno=" & idAnno & " And Utente='" & Utente.Replace("'", "''") & "'"
					Sql = "SELECT UtentiMobile.idAnno, idUtente, Utente, Cognome, Nome, " &
						"Password, EMail, Categorie.idCategoria As idCat1, idTipologia, Categorie.Descrizione As Descr1 " &
						"FROM (UtentiMobile " &
						"Left Join Categorie On UtentiMobile.idCategoria=Categorie.idCategoria And UtentiMobile.idAnno=Categorie.idAnno) " &
						"Where Utente='" & Utente.Replace("'", "''") & "' And UtentiMobile.idAnno=" & idAnno
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							If Password <> Rec("Password").Value.ToString Then
								Ritorno = StringaErrore & " Password non valida"
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					' Sql = "SELECT * FROM UtentiMobile Where idAnno=" & idAnno & " And idUtente=" & idUtente
					Sql = "SELECT UtentiMobile.*, Categorie.Descrizione " &
						"From UtentiMobile LEFT Join Categorie On (UtentiMobile.idCategoria = Categorie.idCategoria) And (UtentiMobile.idAnno = Categorie.idAnno) " &
						"Where idUtente = " & idUtente
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
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
	Public Function RitornaUtenteLocaleDaID(idUtente As String) As String
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

				Try
					' Sql = "SELECT * FROM UtentiMobile Where idAnno=" & idAnno & " And idUtente=" & idUtente
					Sql = "SELECT Utenti.* " &
						"From Utenti " &
						"Where idUtente = " & idUtente
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
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
									"-1;" &
									Rec("idTipologia").Value & ";" &
									";" &
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					Sql = "SELECT UtentiMobile.idAnno, UtentiMobile.idUtente, UtentiMobile.Utente, UtentiMobile.Cognome, UtentiMobile.Nome, UtentiMobile.EMail, Categorie.Descrizione As Categoria, " &
						"UtentiMobile.idTipologia, UtentiMobile.Password, Categorie.idCategoria " &
						"FROM (UtentiMobile LEFT JOIN Categorie ON UtentiMobile.idCategoria = Categorie.idCategoria And UtentiMobile.idAnno = Categorie.idAnno) " &
						"Where UtentiMobile.idAnno=" & idAnno & " Order By 2,1;"
					' "Where UtentiMobile.idAnno=" & idAnno & " Order By 2,1;"
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun utente rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Dim Rec2 As Object

								Sql = " Select * From AnnoAttualeUtenti Where idUtente=" & Rec("idUtente").Value
								Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
								Dim AnnoUtente As Integer = Rec2("idAnno").Value
								Rec2.Close()

								Sql = " Select * From Anni Where idAnno=" & AnnoUtente.ToString
								Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
								Dim NomeSquadra As String = Rec2("NomeSquadra").Value
								Rec2.Close()

								Ritorno &= "0;" &
									Rec("idUtente").Value & ";" &
									Rec("Utente").Value & ";" &
									Rec("Cognome").Value & ";" &
									Rec("Nome").Value & ";" &
									Rec("EMail").Value & ";" &
									NomeSquadra & ";" &
									Rec("idTipologia").Value & ";" &
									Rec("Password").Value & ";" &
									Rec("idCategoria").Value & ";" &
									Rec("Categoria").Value & ";" &
									"§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()

						'Ritorno &= "£"

						'Sql = "Select * From Categorie Where idAnno=" & idAnno & " And Eliminato = 'N' Order By Ordinamento"
						'Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
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
	Public Function EliminaUtente(Squadra As String, ByVal idAnno As String, idUtente As String) As String
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

				Sql = "Delete From UtentiMobile Where idUtente=" & idUtente
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaUtente(Squadra As String, ByVal idAnno As String, Utente As String, Cognome As String, Nome As String, EMail As String,
								Password As String, idCategoria As String, idTipologia As String) As String
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
				Dim idUtente As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "SELECT * FROM UtentiMobile Where Utente='" & Utente.Trim.ToUpper & "' And idAnno=" & idAnno
						Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								If TipoDB = "SQLSERVER" Then
									Sql = "SELECT IsNull(Max(idUtente),0)+1 FROM UtentiMobile" ' Where idAnno=" & idAnno
								Else
									Sql = "SELECT Coalesce(Max(idUtente),0)+1 FROM UtentiMobile" ' Where idAnno=" & idAnno
								End If
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec.Eof() Then
										Ritorno = StringaErrore & " Nessun utente rilevato"
									Else
										'If Rec(0).Value Is DBNull.Value Then
										'	idUtente = "1"
										'Else
										idUtente = Rec(0).Value.ToString
										'End If
									End If
									Rec.Close()
								End If

								If idUtente <> "" Then
									Sql = "Insert Into Utenti Values (" &
										" " & idAnno & ", " &
										" " & idUtente & ", " &
										"'" & Utente.Replace("'", "''") & "', " &
										"'" & Cognome.Replace("'", "''") & "', " &
										"'" & Nome.Replace("'", "''") & "', " &
										"'" & Password.Replace("'", "''") & "', " &
										"'" & EMail.Replace("'", "''") & "', " &
										" " & idCategoria & ", " &
										" " & idTipologia & ", " &
										"'" & stringaWidgets & "' " &
										")"
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If

									Try
										Sql = "Delete From AnnoAttualeUtenti Where idUtente=" & idUtente
										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If

										If Ok Then
											Sql = "Insert Into AnnoAttualeUtenti Values (" & idUtente & ", " & idAnno & ")"
											Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
											End If
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & " " & ex.Message
										Ok = False
									End Try

									If Ritorno = "*" Then Ritorno = idUtente
								Else
									Ritorno = StringaErrore & " Problemi nel rilevamento dell'ID Utente"
									Ok = False
								End If
							Else
								Ritorno = StringaErrore & " Utente già esistente"
								Ok = False
							End If
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
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
	Public Function ModificaUtente(Squadra As String, ByVal idAnno As String, Utente As String, Cognome As String, Nome As String, EMail As String,
								Password As String, idCategoria As String, idTipologia As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec as object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					' Sql = "Delete From UtentiMobile Where idAnno=" & idAnno & " And idUtente=" & idUtente
					Sql = "Delete From UtentiMobile Where idUtente=" & idUtente
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						Try
							Sql = "Insert Into UtentiMobile Values (" &
								"" & idAnno & ", " &
								"" & idUtente & ", " &
								"'" & Utente.Replace("'", "''") & "', " &
								"'" & Cognome.Replace("'", "''") & "', " &
								"'" & Nome.Replace("'", "''") & "', " &
								"'" & Password.Replace("'", "''") & "', " &
								"'" & EMail.Replace("'", "''") & "', " &
								" " & idCategoria & ", " &
								"" & idTipologia & ")"
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If

							If Ok Then
								Try
									Sql = "Delete From AnnoAttualeUtenti Where idUtente=" & idUtente
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If

									If Ok Then
										Sql = "Insert Into AnnoAttualeUtenti Values (" & idUtente & ", " & idAnno & ")"
										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If

									End If
								Catch ex As Exception
									Ritorno = StringaErrore & " " & ex.Message
									Ok = False
								End Try
							End If

							If Ritorno = "*" Then Ritorno = idUtente
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
						End Try
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
End Class