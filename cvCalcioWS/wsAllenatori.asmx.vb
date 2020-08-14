Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_all.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsAllenatori
    Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaNuovoID(Squadra As String, ByVal idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idAllenatore As String = "-1"

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

				Sql = "SELECT Max(idAllenatore)+1 FROM Allenatori Where idAnno=" & idAnno
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec(0).Value Is DBNull.Value Then
						idAllenatore = "1"
					Else
						idAllenatore = Rec(0).Value.ToString
					End If
				End If
				Rec.Close()
			End If
		End If

		Return idAllenatore
	End Function

	<WebMethod()>
	Public Function SalvaAllenatore(Squadra As String, idAnno As String, idCategoria As String, idAllenatore As String,
									Cognome As String, Nome As String, EMail As String, Telefono As String, idTipologia As String,
									TipologiaOperazione As String) As String
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
				Dim idAll As Integer = -1
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					If idAllenatore = "-1" Then
						Try
							Sql = "SELECT Max(idAllenatore)+1 FROM Allenatori Where idAnno=" & idAnno
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec(0).Value Is DBNull.Value Then
									idAll = 1
								Else
									idAll = Rec(0).Value
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
						End Try
					Else
						idAll = idAllenatore
						Sql = "Delete from Allenatori Where idAnno=" & idAnno & " And idAllenatore=" & idAll
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					End If

					If Ok Then
						Sql = "Insert Into Allenatori Values (" &
							" " & idAnno & ", " &
							" " & idCategoria & ", " &
							" " & idAll & ", " &
							"'" & Cognome.Replace("'", "''") & "', " &
							"'" & Nome.Replace("'", "''") & "', " &
							"'" & EMail.Replace("'", "''") & "', " &
							"'" & Telefono.Replace("'", "''") & "', " &
							"'N', " &
							" " & idTipologia & " " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Not Ritorno.Contains(StringaErrore) Then
							If TipologiaOperazione = "INSERIMENTO" Then
								' Aggiunge Utente
								Dim maxGenitore As Integer = -1

								Sql = "Select Max(idUtente) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec(0).Value Is DBNull.Value Then
										maxGenitore = 1
									Else
										maxGenitore = Rec(0).Value
									End If
								End If

								Dim s() As String = Squadra.Split("_")
								Dim idSquadra As Integer = Val(s(1))
								Dim chiave As String = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvZz0123456789!$%/()=?^"
								Dim rnd1 As New Random()
								Dim nuovaPass As String = ""

								For i As Integer = 1 To 7
									Dim c As Integer = rnd1.Next(chiave.Length - 1) + 1
									nuovaPass &= Mid(chiave, c, 1)
								Next

								Dim wrapper As New CryptEncrypt("WPippoBaudo227!")
								Dim nuovaPassCrypt As String = wrapper.EncryptData(nuovaPass)

								Sql = "Insert Into [Generale].[dbo].[Utenti] Values (" &
									" " & idAnno & ", " &
									" " & maxGenitore & ", " &
									"'" & EMail.Replace("'", "''") & "', " &
									"'" & Cognome.Replace("'", "''") & "', " &
									"'" & Nome.Replace("'", "''") & "', " &
									"'" & nuovaPassCrypt.Replace("'", "''") & "', " &
									"'" & EMail.Replace("'", "''") & "', " &
									"-1, " &
									"3, " &
									" " & idSquadra & ", " &
									"1, " &
									"'" & Telefono & "', " &
									"'N', " &
									"-1, " &
									"'N' " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								Else
									Dim m As New mail
									Dim Oggetto As String = "Nuovo utente inCalcio"
									Dim Body As String = ""
									Body &= "E' stato creato l'utente '" & Cognome.ToUpper & " " & Nome.ToUpper & "'. <br />"
									Body &= "Per accedere al sito sarà possibile digitare la mail rilasciata alla segreteria in fase di iscrizione: " & EMail & "<br />"
									Body &= "La password valida per il solo primo accesso è: " & nuovaPass & "<br /><br />"
									Dim ChiScrive As String = "notifiche@incalcio.cloud"

									Ritorno = m.SendEmail("", Oggetto, Body, EMail)
								End If
							End If
						End If
					End If

					If Not Ritorno.Contains(StringaErrore) Then
						Sql = "commit"
						Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
					Else
						Sql = "rollback"
						Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
					End If
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
	Public Function RitornaAllenatoriCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String) As String
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

				Dim Altro As String = ""

				If idCategoria <> "-1" Then
					Altro = "And A.idCategoria=" & idCategoria
				End If
				Try
					Sql = "SELECT A.*, B.Descrizione FROM Allenatori A " &
						"Left Join Categorie B On A.idAnno = B.idAnno And A.idCategoria = B.idCategoria " &
						"Where A.idAnno=" & idAnno & " " & Altro & " And A.Eliminato='N' Order By Cognome, Nome"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun allenatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idAllenatore").Value & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("EMail").Value.ToString.Trim & ";" &
									Rec("Telefono").Value.ToString.Trim & ";" &
									Rec("idCategoria").Value.ToString.Trim & ";" &
									Rec("Descrizione").Value.ToString.Trim & ";" &
									Rec("idTipologia").Value.ToString.Trim & ";" &
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
	Public Function EliminaAllenatore(Squadra As String, ByVal idAnno As String, idAllenatore As String) As String
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
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")

				Sql = "Select * From Allenatori Where idAnno=" & idAnno & " And idAllenatore=" & idAllenatore
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessun allenatore rilevato"
					Else
						Dim EMail As String = Rec("EMail").Value
						Rec.Close()

						Try
							Sql = "Update Allenatori Set Eliminato='S' Where idAnno=" & idAnno & " And idAllenatore=" & idAllenatore
							Ritorno = EsegueSql(Conn, Sql, Connessione)
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
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If
						End If
					End If
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function
End Class