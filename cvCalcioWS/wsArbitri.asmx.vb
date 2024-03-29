﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_arb.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsArbitri
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaNuovoID(Squadra As String, ByVal idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idArbitro As String = "-1"

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
					Sql = "SELECT IsNull(Max(idArbitro),0)+1 FROM Arbitri Where idAnno=" & idAnno
				Else
					Sql = "SELECT Coalesce(Max(idArbitro),0)+1 FROM Arbitri Where idAnno=" & idAnno
				End If
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	idArbitro = "1"
					'Else
					idArbitro = Rec(0).Value.ToString
					'End If
				End If
				Rec.Close()
			End If
		End If

		Return idArbitro
	End Function

	<WebMethod()>
	Public Function SalvaArbitro(Squadra As String, idAnno As String, idCategoria As String, idArbitro As String,
								 Cognome As String, Nome As String, EMail As String, Telefono As String) As String
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
					If idArbitro = "-1" Then
						Try
							Sql = "SELECT " & IIf(TipoDB = "SQLSERVER", "IsNull(Max(idArbitro),0)+1", "Coalesce(Max(idArbitro),0)+1") & " FROM Arbitri Where idAnno=" & idAnno
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
							Ok = False
							Ritorno = StringaErrore & " " & ex.Message
						End Try
					Else
						idDir = idArbitro
						Sql = "Delete from Arbitri Where idAnno=" & idAnno & " And idArbitro=" & idDir
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						Sql = "Insert Into Arbitri Values (" &
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
					End If

					If Not Ritorno.Contains(StringaErrore) Then
						Ritorno = "*"

						Sql = "commit"
						Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					Else
						Sql = "rollback"
						Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					End If
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
	Public Function RitornaArbitri(Squadra As String, ByVal idAnno As String, idCategoria As String) As String
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
					' Sql = "SELECT * FROM Arbitri Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Eliminato='N' Order By Cognome, Nome"
					Sql = "SELECT * FROM Arbitri Where idAnno=" & idAnno & " And Eliminato='N' Order By Cognome, Nome"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun Arbitro rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("idArbitro").Value & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("EMail").Value.ToString.Trim & ";" &
									Rec("Telefono").Value.ToString.Trim & ";" &
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
	Public Function EliminaArbitro(Squadra As String, ByVal idAnno As String, idArbitro As String) As String
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

				Try
					Sql = "Update Arbitri Set Eliminato='S' Where idAnno=" & idAnno & " And idArbitro=" & idArbitro
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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