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
					Sql = "SELECT A.*, B.descrizione From PermessoVisual A , Visualizza B
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
										   Rec("descrizione").Value.ToString & "§"

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
					Sql = "SELECT * From Visualizza 
							Order By IDfunzione"
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

				Try
					Sql = "Delete PermessoVisual Where IDutente=" & IDutente & " AND progressivo=" & progressivo
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
						'Rec.Close()
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
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ritorno.Contains(StringaErrore) Then
					Dim Ritorno2 As String

					Sql = "Delete From PermessoVisual Where IDutente=" & IDutente
					Ritorno2 = EsegueSql(Conn, Sql, Connessione)

				End If

				'Conn.Close()
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

				Sql = "Delete From PermessoVisual Where IDutente=" & IDutente & " AND progressivo=" & progressivo
				Ritorno = EsegueSql(Conn, Sql, Connessione)
				'End If

				Sql = "Insert Into PermessoVisual Values (" &
							" " & SistemaNumero(IDutente) & "," &
							" " & SistemaNumero(progressivo) & "," &
							" " & SistemaNumero(permesso) & " " &
							")"


				Ritorno = EsegueSql(Conn, Sql, Connessione)

				''Conn.Close()
			End If
		End If

		Return Ritorno
	End Function


End Class