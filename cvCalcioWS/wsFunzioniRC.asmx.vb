﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://wsFunzioniRC.PAndE.it/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsFunzioniRC
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaFunzioni(Squadra As String) As String
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
					' where IDfunzione = " & IDfunzione & " 
					Sql = "SELECT * From Funzione 
							Order By IDfunzione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna funzione ritornata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								' Ritorno &= Rec("IDfunzione").Value.ToString & ";" & Rec("descrizione").Value.ToString & ";" & Rec("tipo_numero").Value.ToString & "§"
								Ritorno &= Rec("IDfunzione").Value.ToString & ";" & Rec("descrizione").Value.ToString & "§"

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
	Public Function EliminaFunzioni(Squadra As String, IDfunzione As Integer) As String
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
					Sql = "Delete Funzione Where IDfunzione=" & IDfunzione
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
	Public Function InserisciFunzione(Squadra As String, IDfunzione As Integer, descrizione As String) As String
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
				Dim ProgFunz As Integer = -1

				Try
					Sql = "SELECT Max(IDfunzione)+1 FROM Funzione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec(0).Value Is DBNull.Value Then
							ProgFunz = 1
						Else
							ProgFunz = Rec(0).Value
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
							Sql = "Insert Into Funzione Values (" &
								" " & IDfunzione & "," &
								"'" & descrizione.Replace("'", "''") & "' " &
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

					Sql = "Delete From Funzione Where IDfunzione=" & IDfunzione
					Ritorno2 = EsegueSql(Conn, Sql, Connessione)

				End If

				'Conn.Close()
			End If
		End If

		Return Ritorno
	End Function



	<WebMethod()>
	Public Function ModificaFunzione(Squadra As String, IDfunzione As Integer, descrizione As String) As String
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
				Dim ProgFunz As Integer = -1

				If IDfunzione = -1 Then
					Try
						Sql = "SELECT Max(IDfunzione)+1 FROM Funzione"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec(0).Value Is DBNull.Value Then
								ProgFunz = 1
							Else
								ProgFunz = Rec(0).Value
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				Else
					ProgFunz = IDfunzione
					Sql = "Delete From Funzione Where IDfunzione=" & IDfunzione
					Ritorno = EsegueSql(Conn, Sql, Connessione)
				End If

				Sql = "Insert Into Funzione Values (" &
								" " & IDfunzione & "," &
								"'" & descrizione.Replace("'", "''") & "' " &
								")"

				Ritorno = EsegueSql(Conn, Sql, Connessione)

				''Conn.Close()
			End If
		End If

		Return Ritorno
	End Function


End Class