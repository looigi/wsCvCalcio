﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvTaglie.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsTaglie
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaTaglie(Squadra As String) As String
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
					Sql = "SELECT * FROM Taglie Where Elminato='N' Order By Descrizione"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessuna taglia rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("idTaglia").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

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
	Public Function EliminaTaglia(Squadra As String, ByVal idTaglia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Giocata As String = ""

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Sql = "Select * From Giocatori Where idTaglia=" & idTaglia
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof() Then
							Ritorno = StringaErrore & " La taglia è utilizzata"
							Ok = False
						End If
						Rec.Close()
					End If

					If Ok Then
						Try
							Sql = "Update Taglie Set Elminato='S' " &
							"Where idTaglia=" & idTaglia
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaTaglia(Squadra As String, ByVal idTaglia As String, Descrizione As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Giocata As String = ""

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Update Taglie Set Descrizione='" & Descrizione.Replace("'", "''") & "' " &
							"Where idTaglia=" & idTaglia
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InserisceTaglia(Squadra As String, Descrizione As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Giocata As String = ""

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

				Dim idTaglia As Integer = -1

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						If TipoDB = "SQLSERVER" Then
							Sql = "SELECT IsNull(Max(idTaglia),0)+1 FROM Taglie"
						Else
							Sql = "SELECT Coalesce(Max(idTaglia),0)+1 FROM Taglie"
						End If
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							'If Rec(0).Value Is DBNull.Value Then
							'	idTaglia = 1
							'Else
							idTaglia = Rec(0).Value
							'End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						Try
							Sql = "Insert Into Taglie Values (" & idTaglia & ", '" & Descrizione.Replace("'", "''") & "', 'N')"
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							Else
								Sql = "Select * From Taglie Where Descrizione='" & Descrizione.Replace("'", "''") & "'"
								Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									Ritorno = Rec("idTaglia").Value
								End If
							End If

						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

End Class