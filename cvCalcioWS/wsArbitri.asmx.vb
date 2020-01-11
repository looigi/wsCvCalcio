Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_arb.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsArbitri
    Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function SalvaArbitro(Squadra As String, idAnno As String, idCategoria As String, idArbitro As String,
								 Cognome As String, Nome As String, EMail As String, Telefono As String) As String
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
				Dim idDir As Integer = -1

				If idArbitro = "-1" Then
					Try
						Sql = "SELECT Max(idArbitro)+1 FROM Arbitri Where idAnno=" & idAnno
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec(0).Value Is DBNull.Value Then
								idDir = 1
							Else
								idDir = Rec(0).Value
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				Else
					idDir = idArbitro
					Sql = "delete from Arbitri Where idAnno=" & idAnno & " And idArbitro=" & idDir
					Ritorno = EsegueSql(Conn, Sql, Connessione)
				End If

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
				Ritorno = EsegueSql(Conn, Sql, Connessione)

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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Try
					' Sql = "SELECT * FROM Arbitri Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Eliminato='N' Order By Cognome, Nome"
					Sql = "SELECT * FROM Arbitri Where idAnno=" & idAnno & " And Eliminato='N' Order By Cognome, Nome"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun Arbitro rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Try
					Sql = "Update Arbitri Set Eliminato='S' Where idAnno=" & idAnno & " And idArbitro=" & idArbitro
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