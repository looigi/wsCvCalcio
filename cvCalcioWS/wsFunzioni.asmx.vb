﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Web.Hosting
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://wsFunzioni.PAndE.it/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsFunzioni
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaFunzionalita() As String
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
					' where IDfunzione = " & IDfunzione & " 
					Sql = "SELECT * From Permessi_Lista Where Eliminato='N' " &
						"Order By Descrizione"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = "" ' StringaErrore & " Nessuna funzionalità ritornata"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("idPermesso").Value.ToString & ";" & Rec("Descrizione").Value.ToString & ";" & Rec("NomePerCodice").Value.ToString & ";§"

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
	Public Function InserisciFunzionalita(Descrizione As String, Codice As String) As String
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
				Dim Successivo As Integer = -1
				Dim Ok As Boolean = True

				Try
					If TipoDB = "SQLSERVER" Then
						Sql = "SELECT IsNull(Max(idPermesso),0)+1 FROM Permessi_Lista"
					Else
						Sql = "SELECT Coalesce(Max(idPermesso),0)+1 FROM Permessi_Lista"
					End If
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						'If Rec(0).Value Is DBNull.Value Then
						'	Successivo = 1
						'Else
						Successivo = Rec(0).Value
						'End If
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				If Ok Then
					Try
						Sql = "Insert Into Permessi_Lista Values (" &
							" " & Successivo & ", " &
							"'" & Descrizione.Replace("'", "''") & "' ," &
							"'" & Codice.Replace("'", "''") & "' ," &
							"'N' " &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaFunzionalita(idPermesso As String) As String
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
				Dim Successivo As Integer = -1

				Try
					Sql = "Update Permessi_Lista Set Eliminato='S' Where idPermesso=" & idPermesso
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaFunzionalita(idPermesso As String, Descrizione As String, Codice As String) As String
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
					Sql = "Update Permessi_Lista Set Descrizione='" & Descrizione.Replace("'", "''") & "', NomePerCodice='" & Codice.Replace("'", "''") & "' Where idPermesso=" & idPermesso
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaTutteLeFunzioni(idTipologia As String) As String
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
					' where IDfunzione = " & IDfunzione & " 
					Sql = "SELECT * From Permessi_Lista Where idPermesso Not In (Select idPermesso From Permessi_Composizione Where idTipologia = " & idTipologia & ") And Eliminato='N' " &
						"Order By Descrizione"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = "" ' StringaErrore & " Nessuna funzionalità ritornata"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								' Ritorno &= Rec("IDfunzione").Value.ToString & ";" & Rec("descrizione").Value.ToString & ";" & Rec("tipo_numero").Value.ToString & "§"
								Ritorno &= Rec("idPermesso").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

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
	Public Function RitornaFunzioni(idTipologia As String) As String
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
					' where IDfunzione = " & IDfunzione & " 
					Sql = "SELECT A.idPermesso, B.Descrizione From Permessi_Composizione A " &
						"Left Join Permessi_Lista B On A.idPermesso = B.idPermesso " &
						"Where A.idTipologia = " & idTipologia & " " &
						"Order By B.Descrizione"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessuna funzionalità ritornata"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								' Ritorno &= Rec("IDfunzione").Value.ToString & ";" & Rec("descrizione").Value.ToString & ";" & Rec("tipo_numero").Value.ToString & "§"
								Ritorno &= Rec("idPermesso").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

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
	Public Function EliminaFunzioni(IDfunzione As Integer, idTipologia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Delete From Permessi_Composizione Where idPermesso=" & IDfunzione & " And idTipologia=" & idTipologia
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
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
	Public Function InserisciFunzione(IDfunzione As Integer, idTipologia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim Rec As Object
				Dim ProgFunz As Integer = -1

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						If TipoDB = "SQLSERVER" Then
							Sql = "SELECT IsNull(Max(Progressivo),0)+1 FROM Permessi_Composizione Where idTipologia=" & idTipologia
						Else
							Sql = "SELECT Coalesce(Max(Progressivo),0)+1 FROM Permessi_Composizione Where idTipologia=" & idTipologia
						End If
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							'If Rec(0).Value Is DBNull.Value Then
							'	ProgFunz = 1
							'Else
							ProgFunz = Rec(0).Value
							'End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						Try
							If Not Ritorno.Contains(StringaErrore) Then
								Sql = "Insert Into Permessi_Composizione Values (" &
									" " & idTipologia & ", " &
									" " & ProgFunz & ", " &
									" " & IDfunzione & " " &
									")"
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

					If Ritorno.Contains(StringaErrore) Then
						Dim Ritorno2 As String

						Sql = "Delete From Permessi_Composizione Where idPermesso=" & IDfunzione & " And idTipologia=" & idTipologia
						Ritorno2 = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					End If
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

	'<WebMethod()>
	'Public Function ModificaFunzione(Squadra As String, IDfunzione As Integer, descrizione As String, idTipologia As String) As String
	'	Dim Ritorno As String = ""
	'	Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

	'	If Connessione = "" Then
	'		Ritorno = ErroreConnessioneNonValida
	'	Else
	'		Dim Conn As Object = New clsGestioneDB(Squadra)

	'		If TypeOf (Conn) Is String Then
	'			Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
	'		Else
	'			Dim Rec as object
	'			Dim Sql As String = ""
	'			Dim ProgFunz As Integer = -1
	'			Dim Ok As Boolean = True

	'			Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
	'			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

	'			If Not Ritorno.Contains(StringaErrore) Then
	'				If IDfunzione = -1 Then
	'					Try
	'						Sql = "SELECT Max(IDfunzione)+1 FROM Funzione"
	'						Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
	'						If TypeOf (Rec) Is String Then
	'							Ritorno = Rec
	'						Else
	'							If Rec(0).Value Is DBNull.Value Then
	'								ProgFunz = 1
	'							Else
	'								ProgFunz = Rec(0).Value
	'							End If
	'							Rec.Close()
	'						End If
	'					Catch ex As Exception
	'						Ritorno = StringaErrore & " " & ex.Message
	'						Ok = False
	'					End Try
	'				Else
	'					ProgFunz = IDfunzione
	'					Sql = "Delete From Funzione Where IDfunzione=" & IDfunzione
	'					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'					Ok = False
	'				End If

	'				If Ok Then
	'					Sql = "Insert Into Funzione Values (" &
	'								" " & IDfunzione & "," &
	'								"'" & descrizione.Replace("'", "''") & "' " &
	'								")"

	'					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'					If Ritorno.Contains(StringaErrore) Then
	'						Ok = False
	'					End If
	'				End If
	'			Else
	'				Ok = False
	'			End If

	'			If Ok Then
	'				Sql = "commit"
	'				Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'			Else
	'				Sql = "rollback"
	'				Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'			End If

	'			Conn.Close()
	'		End If
	'	End If

	'	Return Ritorno
	'End Function
End Class