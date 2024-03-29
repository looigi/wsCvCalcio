﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.Odbc
Imports System.Web.UI.WebControls.Adapters
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_dir.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsDirigenti
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaNuovoID(Squadra As String, ByVal idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idDirigente As String = "-1"

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

				If TipoDB = "SQLSERVER" Then
					Sql = "SELECT IsNull(Max(idDirigente),0)+1 FROM Dirigenti Where idAnno=" & idAnno
				Else
					Sql = "SELECT Coalesce(Max(idDirigente),0)+1 FROM Dirigenti Where idAnno=" & idAnno
				End If
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	idDirigente = "1"
					'Else
					idDirigente = Rec(0).Value.ToString
					'End If
				End If
				Rec.Close()
			End If
		End If

		Return idDirigente
	End Function

	<WebMethod()>
	Public Function SalvaDirigente(Squadra As String, idAnno As String, idCategoria As String, idDirigente As String,
									Cognome As String, Nome As String, EMail As String, Telefono As String,
									TipologiaOperazione As String, Tendina As String, Mittente As String) As String
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
				Dim idDir As Integer = -1
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					If idDirigente = "-1" Then
						Try
							If TipoDB = "SQLSERVER" Then
								Sql = "SELECT IsNull(Max(idDirigente),0)+1 FROM Dirigenti Where idAnno=" & idAnno
							Else
								Sql = "SELECT Coalesce(Max(idDirigente),0)+1 FROM Dirigenti Where idAnno=" & idAnno
							End If
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								'If Rec(0).Value Is DBNull.Value Then
								'	idDir = 1
								'Else
								idDir = Rec(0).Value
								'End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					Else
						idDir = idDirigente
						Sql = "Delete From Dirigenti Where idAnno=" & idAnno & " And idDirigente=" & idDir
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok = True Then
						Sql = "Insert Into Dirigenti Values (" &
							" " & idAnno & ", " &
							" " & idCategoria & ", " &
							" " & idDir & ", " &
							"'" & Cognome.Replace("'", "''") & "', " &
							"'" & Nome.Replace("'", "''") & "', " &
							"'" & EMail.Replace("'", "''") & "', " &
							"'" & Telefono.Replace("'", "''") & "', " &
							"'N' " &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						Else
							If Not Ritorno.Contains(StringaErrore) Then
								Sql = "Delete From DirigentiCategorie Where idUtente = " & idDir
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							End If
							If Not Ritorno.Contains(StringaErrore) Then
								Sql = "Insert Into DirigentiCategorie Values (" & idDir & ", 1, " & idCategoria & ")"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							End If

							If TipologiaOperazione = "INSERIMENTO" Then
								' Aggiunge Utente
								Dim idGenitore As Integer = -1

								If Tendina = "N" Then
									If TipoDB = "SQLSERVER" Then
										Sql = "Select IsNull(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
									Else
										Sql = "Select Coalesce(Max(idUtente),0) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
									End If
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										'If Rec(0).Value Is DBNull.Value Then
										'	idGenitore = 1
										'Else
										idGenitore = Rec(0).Value
										'End If
									End If
								Else
									Sql = "Select * From [Generale].[dbo].[Utenti] Where EMail='" & EMail.Replace("'", "''") & "'"
									Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										If Rec.Eof() Then
											Ritorno = StringaErrore & " Nessun utente rilevato"
											Ok = False
										Else
											idGenitore = Rec("idUtente").Value
										End If
									End If
								End If

								Dim pass As String = generaPassRandom()
								Dim nuovaPass() = pass.Split(";")

								If Tendina = "S" Then
									Sql = "Update [Generale].[dbo].[Utenti] Set idTipologia=4 Where idUtente=" & idGenitore
								Else
									Dim s() As String = Squadra.Split("_")
									Dim idSquadra As Integer = Val(s(1))

									Sql = "Insert Into [Generale].[dbo].[Utenti] Values (" &
										" " & idAnno & ", " &
										" " & idGenitore & ", " &
										"'" & EMail.Replace("'", "''") & "', " &
										"'" & Cognome.Replace("'", "''") & "', " &
										"'" & Nome.Replace("'", "''") & "', " &
										"'" & nuovaPass(1).Replace("'", "''") & "', " &
										"'" & EMail.Replace("'", "''") & "', " &
										"-1, " &
										"8, " &
										" " & idSquadra & ", " &
										"1, " &
										"'" & Telefono & "', " &
										"'N', " &
										"-1, " &
										"'N', " &
										"'" & stringaWidgets & "' " &
										")"
								End If
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								Else
									Ritorno = CreaPermessiDiBase(Conn, Connessione, idGenitore, idCategoria, Tendina)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									Else
										If Tendina = "N" Then
											Dim m As New mail
											Dim Oggetto As String = "Nuovo dirigente inCalcio"
											Dim Body As String = ""
											Body &= "E' stato creato il dirigente '" & Cognome.ToUpper & " " & Nome.ToUpper & "'. <br />"
											Body &= "Per accedere al sito sarà possibile digitare la mail rilasciata alla segreteria in fase di iscrizione: " & EMail & "<br />"
											Body &= "La password valida per il solo primo accesso è: " & nuovaPass(0) & "<br /><br />"
											Dim ChiScrive As String = "notifiche@incalcio.cloud"

											Ritorno = m.SendEmail(Server.MapPath("."), Squadra, Mittente, Oggetto, Body, EMail, {""})
										End If
									End If
								End If
							End If
						End If
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

	Private Function CreaPermessiDiBase(Conn As Object, Connessione As String, idGenitore As Integer, idCategoria As Integer, Tendina As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String = ""
		Dim Ok As Boolean = True
		Dim Rec As Object

		If Tendina = "S" Then
			Sql = "Delete From DirigentiCategorie Where idUtente=" & idGenitore
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

			Sql = "Delete From PermessiUtente Where idUtente=" & idGenitore
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
		End If

		Sql = "Insert Into DirigentiCategorie Values (" &
				" " & idGenitore & ", " &
				"1, " &
				" " & idCategoria & " " &
				")"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
		If Ritorno.Contains(StringaErrore) Then
			Ok = False
		Else
			Sql = "Select * From [Generale].[dbo].[Permessi_LIsta] Where NomePerCodice In ('HOME', 'CALENDARIO', 'ROSE', 'DIRIGENTI', 'ALLENAMENTI', 'PARTITE', 'CAMPIONATO', 'STATISTICHE', 'CONTATTI', 'CONVOCAZIONI')"
			Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				Dim idPermesso As New List(Of Integer)

				Do Until Rec.Eof()
					idPermesso.Add(Rec("idPermesso").Value)

					Rec.MoveNext()
				Loop
				Rec.Close()

				Dim Progressivo As Integer = 1

				For Each id As Integer In idPermesso
					Sql = "Insert Into PermessiUtente Values (" &
						" " & idGenitore & ", " &
						" " & Progressivo & ", " &
						" " & id & " " &
						")"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					Progressivo += 1
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
						Exit For
					End If
				Next
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaDirigentiCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String) As String
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

				Dim Altro As String = ""

				If idCategoria <> "-1" Then
					Altro = "And B.idCategoria=" & idCategoria
				End If
				Try
					' Sql = "SELECT * FROM Dirigenti Where idAnno=" & idAnno & " " & Altro & " And Eliminato='N' Order By Cognome, Nome"
					'Sql = "SELECT A.*, B.Descrizione FROM Dirigenti A " &
					'	"Left Join Categorie B On A.idAnno = B.idAnno And A.idCategoria = B.idCategoria " &
					'	"Where A.idAnno=" & idAnno & " " & Altro & " And A.Eliminato='N' Order By Cognome, Nome"

					Sql = "Select * From DirigentiCategorie A " &
						"Join Dirigenti B On A.idUtente = B.idDirigente " &
						"Left Join Categorie C On A.idCategoria = C.idCategoria " &
						"WHere B.idAnno=" & idAnno & " " & Altro & " And B.Eliminato = 'N' Order By Cognome, Nome"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun Dirigente rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("idDirigente").Value & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("EMail").Value.ToString.Trim & ";" &
									Rec("Telefono").Value.ToString.Trim & ";" &
									Rec("idCategoria").Value.ToString.Trim & ";" &
									Rec("Descrizione").Value.ToString.Trim & ";" &
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
	Public Function EliminaDirigente(Squadra As String, ByVal idAnno As String, idDirigente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Dim Rec As Object

					Sql = "Select * From Dirigenti Where idAnno=" & idAnno & " And idDirigente=" & idDirigente
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun Dirigente rilevato"
						Else
							Dim EMail As String = Rec("EMail").Value
							Rec.Close()

							Try
								Sql = "Update Dirigenti Set Eliminato='S' Where idAnno=" & idAnno & " And idDirigente=" & idDirigente
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try

							If Ok Then
								Sql = "Update [Generale].[dbo].[Utenti] Set Eliminato='S' " &
									"Where Utente='" & EMail.Replace("'", "''") & "'"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							End If
						End If
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