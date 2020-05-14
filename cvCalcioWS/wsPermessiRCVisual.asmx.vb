Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://wsPermessiRCVisual.PAndE.it/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsPermessiRCVisual
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaPermessiVisual(Squadra As String, IDutente As Integer) As String
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
					Sql = "SELECT A.*, B.descrizione, B.NomePerCodice From PermessoVisual A , Visualizza B
							where IDutente=" & IDutente & " 
							and   permesso = IDfunzione
							Order By progressivo"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun permesso ritornato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("IDutente").Value.ToString & ";" & Rec("progressivo").Value.ToString & ";" & Rec("permesso").Value.ToString & ";" &
										   Rec("descrizione").Value.ToString & ";" & Rec("NomePerCodice").Value.ToString & "§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					'				Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaTuttiPermessiVisual(Squadra As String) As String
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
					Sql = "SELECT * From Visualizza Where Eliminato = 'N' 
							Order By Descrizione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun permesso ritornato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("IDfunzione").Value.ToString & ";" & Rec("descrizione").Value.ToString & "§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					'				Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaPermessoVisual(Squadra As String, IDutente As Integer, progressivo As Integer) As String
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
				Dim PermessoDaEliminare As String = ""

				Try
					Sql = "Select B.descrizione From PermessoVisual A Left Join " &
						"Visualizza B On A.permesso = B.IDfunzione " &
						"Where IDutente = " & IDutente & " And progressivo = " & progressivo
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ok = False
							Ritorno = StringaErrore & " Permesso non rilevato"
						Else
							PermessoDaEliminare = Rec(0).Value
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				If PermessoDaEliminare.Contains("/") Then
					Dim Campi() As String = PermessoDaEliminare.Split("/")
					Dim Chiave As String = Campi(0).Trim & " /"
					Dim CiSonoAltri As Boolean = False

					Try
						Sql = "Select B.descrizione From PermessoVisual A Left Join " &
							"Visualizza B On A.permesso = B.IDfunzione " &
							"Where IDutente = " & IDutente & " And Progressivo <> " & progressivo
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ok = False
								Ritorno = StringaErrore & " Permessi non rilevato"
							Else
								Do Until Rec.Eof
									If Rec(0).Value.ToString.Contains(Chiave) Then
										CiSonoAltri = True
										Exit Do
									End If

									Rec.MoveNext
								Loop
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Not CiSonoAltri Then
						Dim IdPermessoPadre As Integer = -1

						Try
							Sql = "Select * From Visualizza Where descrizione = '" & Chiave.Replace("/", "").Trim & "'"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ok = False
									Ritorno = StringaErrore & " Permesso padre non rilevato"
								Else
									IdPermessoPadre = Rec("IDfunzione").Value
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try

						If IdPermessoPadre > -1 Then
							Try
								Sql = "Delete PermessoVisual Where IDutente=" & IDutente & " AND permesso = " & IdPermessoPadre
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If
					End If
				End If

				Try
					Sql = "Delete PermessoVisual Where IDutente=" & IDutente & " AND progressivo=" & progressivo
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InserisciPermessoVisual(Squadra As String, IDutente As Integer, progressivo As Integer, permesso As Integer) As String
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
				Dim ProgPerm As Integer = -1

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "SELECT Max(progressivo)+1 FROM PermessoVisual Where IDutente=" & IDutente
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec(0).Value Is DBNull.Value Then
								ProgPerm = 1
							Else
								ProgPerm = Rec(0).Value
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					Dim DescrizionePermesso As String = ""

					Try
						Sql = "SELECT * FROM Visualizza Where idFunzione=" & permesso
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Permesso non trovato"
							Else
								DescrizionePermesso = Rec("descrizione").Value
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Not Ritorno.Contains(StringaErrore) Then
						Try
							If Not Ritorno.Contains(StringaErrore) Then
								Sql = "Insert Into PermessoVisual Values (" &
									" " & SistemaNumero(IDutente) & "," &
									" " & SistemaNumero(ProgPerm) & "," &
									" " & SistemaNumero(permesso) & " " &
									")"

								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If

					If DescrizionePermesso.Contains("/") Then
						Dim lPadre() As String = DescrizionePermesso.Split("/")
						Dim Padre As String = lPadre(0).Trim
						Dim idFunzionePadre As Integer = -1

						Try
							Sql = "SELECT * FROM Visualizza Where descrizione = '" & Padre & "'"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = StringaErrore & " Permesso padre non trovato"
								Else
									idFunzionePadre = Rec("IDfunzione").Value
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try

						Dim DaAggiungere As Boolean = False

						If Not Ritorno.Contains(StringaErrore) Then
							Try
								Sql = "SELECT * FROM PermessoVisual Where IDutente = " & IDutente & " And permesso = " & idFunzionePadre
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec.Eof Then
										DaAggiungere = True
									End If
									Rec.Close()
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If

						If Not Ritorno.Contains(StringaErrore) And DaAggiungere Then
							Try
								If Not Ritorno.Contains(StringaErrore) Then
									ProgPerm += 1
									Sql = "Insert Into PermessoVisual Values (" &
									" " & SistemaNumero(IDutente) & "," &
									" " & SistemaNumero(ProgPerm) & "," &
									" " & SistemaNumero(idFunzionePadre) & " " &
									")"

									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If
					End If

						If Ritorno.Contains(StringaErrore) Then
						Dim Ritorno2 As String

						Sql = "Delete From PermessoVisual Where IDutente=" & IDutente
						Ritorno2 = EsegueSql(Conn, Sql, Connessione)
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
	Public Function ModificaPermessoVisual(Squadra As String, IDutente As Integer, progressivo As Integer, permesso As Integer) As String
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
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Sql = "Delete From PermessoVisual Where IDutente=" & IDutente & " AND progressivo=" & progressivo
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						Sql = "Insert Into PermessoVisual Values (" &
								" " & SistemaNumero(IDutente) & "," &
								" " & SistemaNumero(progressivo) & "," &
								" " & SistemaNumero(permesso) & " " &
								")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
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


End Class