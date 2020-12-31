Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_avv.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsAvversari
    Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaNuovoID(Squadra As String, ByVal idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idAvversario As String = "-1"

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

				Sql = "SELECT Max(idAvversario)+1 FROM Avversari"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec(0).Value Is DBNull.Value Then
						idAvversario = "1"
					Else
						idAvversario = Rec(0).Value.ToString
					End If
					Rec.Close()
				End If
			End If
		End If

		Return idAvversario
	End Function

	<WebMethod()>
	Public Function RitornaAvversari(Squadra As String, ByVal idAnno As String, Ricerca As String) As String
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

				If Ricerca.Trim <> "" Then
					Altro = "And SquadreAvversarie.Descrizione Like '%" & Ricerca & "%' "
				End If
				Try
					Sql = "SELECT SquadreAvversarie.idAvversario, SquadreAvversarie.idCampo, SquadreAvversarie.Descrizione, " &
						"CampiAvversari.Descrizione As Campo, Indirizzo, Lat, Lon, Telefono, Referente, EMail, Fax " &
						"FROM (SquadreAvversarie " &
						"Left Join CampiAvversari On SquadreAvversarie.idCampo=CampiAvversari.idCampo) " &
						"Left Join AvversariCoord On AvversariCoord.idAvversario=SquadreAvversarie.idAvversario " &
						"Where SquadreAvversarie.Eliminato='N' " & Altro & "Order By SquadreAvversarie.Descrizione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " No avversaries found"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idAvversario").Value.ToString & ";" & Rec("idCampo").Value.ToString & ";" & Rec("Descrizione").Value.ToString.Trim & ";" & Rec("Campo").Value.ToString.Trim & ";" &
									Rec("Indirizzo").Value.ToString.Trim & ";" & Rec("Lat").Value.ToString & ";" & Rec("Lon").Value.ToString & ";" &
									Rec("Telefono").Value.ToString & ";" & Rec("Referente").Value.ToString & ";" & Rec("EMail").Value.ToString & ";" &
									Rec("Fax").Value.ToString & ";§"

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
	Public Function SalvaAvversario(Squadra As String, idAnno As String, idAvversario As String, idCampo As String, Avversario As String,
									Campo As String, Indirizzo As String, Coords As String, Telefono As String, Referente As String,
									EMail As String, Fax As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim Ok As Boolean = True

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim idAvv As Integer = -1
				Dim idCam As Integer = -1

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					If idAvversario = "-1" Then
						Try
							Sql = "SELECT Max(idAvversario)+1 FROM SquadreAvversarie"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec(0).Value Is DBNull.Value Then
									idAvv = 1
								Else
									idAvv = Rec(0).Value
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					Else
						idAvv = idAvversario
						Sql = "Delete from SquadreAvversarie Where idAvversario=" & idAvv
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok = True And idCampo = "-1" Then
						Try
							Sql = "SELECT Max(idCampo)+1 FROM CampiAvversari"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec(0).Value Is DBNull.Value Then
									idCam = 1
								Else
									idCam = Rec(0).Value
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					Else
						idCam = idCampo
						Sql = "Delete from CampiAvversari Where idCampo=" & idCam
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						Sql = "Insert Into SquadreAvversarie Values (" &
							" " & idCam & ", " &
							" " & idAvv & ", " &
							"'" & Avversario.Replace("'", "''") & "', " &
							"'N', " &
							"'" & Telefono.Replace("'", "''") & "', " &
							"'" & Referente.Replace("'", "''") & "', " &
							"'" & EMail.Replace("'", "''") & "', " &
							"'" & Fax.Replace("'", "''") & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						Sql = "Insert Into CampiAvversari Values (" &
							" " & idCam & ", " &
							"'" & Campo.Replace("'", "''") & "', " &
							"'" & Indirizzo.Replace("'", "''") & "', " &
							"'N' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						Sql = "Delete From AvversariCoord Where idAvversario=" & idAvv
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						Dim cc() As String = Coords.Split(";")

						Sql = "Insert Into AvversariCoord Values (" &
							" " & idAvv & ", " &
							"'" & cc(0) & "', " &
							"'" & cc(1) & "' " &
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

	<WebMethod()>
	Public Function EliminaAvversario(Squadra As String, ByVal idAnno As String, idAvversario As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), squadra)

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
					Sql = "Update SquadreAvversarie Set Eliminato='S' Where idAvversario=" & idAvversario
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
	Public Function RitornaStatisticheAvversario(Squadra As String, ByVal idAnno As String, idAvversario As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Altro As String = ""

				Try
					Dim IncontriTotali As Integer = EsegueStatistica(Conn, Connessione, "Select Count(*) From Partite Where idAvversario=" & idAvversario & " And Giocata='S'")
					Dim IncontriAnno As Integer = EsegueStatistica(Conn, Connessione, "Select Count(*) From Partite Where idAvversario=" & idAvversario & " And Giocata='S' And idAnno=" & idAnno)

					Dim PartInCasa As Integer = 0
					Dim VittInCasa As Integer = 0
					Dim PareInCasa As Integer = 0
					Dim SconInCasa As Integer = 0
					Dim GfInCasa As Integer = 0
					Dim GSInCasa As Integer = 0

					Dim PartFuoriCasa As Integer = 0
					Dim VittFuoriCasa As Integer = 0
					Dim PareFuoriCasa As Integer = 0
					Dim SconFuoriCasa As Integer = 0
					Dim GfFuoriCasa As Integer = 0
					Dim GSFuoriCasa As Integer = 0

					Dim PartInCasaAnno As Integer = 0
					Dim VittInCasaAnno As Integer = 0
					Dim PareInCasaAnno As Integer = 0
					Dim SconInCasaAnno As Integer = 0
					Dim GfInCasaAnno As Integer = 0
					Dim GSInCasaAnno As Integer = 0

					Dim PartFuoriCasaAnno As Integer = 0
					Dim VittFuoriCasaAnno As Integer = 0
					Dim PareFuoriCasaAnno As Integer = 0
					Dim SconFuoriCasaAnno As Integer = 0
					Dim GfFuoriCasaAnno As Integer = 0
					Dim GSFuoriCasaAnno As Integer = 0

					Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
					Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
					Sql = "SELECT Partite.idAnno, Partite.idPartita, Partite.Casa, Risultati.Risultato, RisultatiAggiuntivi.RisGiochetti,GoalAvvPrimoTempo,GoalAvvSecondoTempo,GoalAvvTerzoTempo " &
						"FROM (Partite " &
						"Left Join Risultati On Partite.idPartita = Risultati.idPartita) " &
						"Left Join RisultatiAggiuntivi On Partite.idPartita = RisultatiAggiuntivi.idPartita " &
						"WHERE Partite.idAvversario=" & idAvversario & " And Giocata='S'"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Dim Anno As Integer = Rec("idAnno").Value
							Dim idPartita As String = Rec("idPartita").Value
							Dim Risultato As String = "" & Rec("Risultato").Value
							Dim RisGiochetti As String = "" & Rec("RisGiochetti").Value
							Dim GoalSegnatiAvv As Integer = 0
							GoalSegnatiAvv = Val("" & Rec("GoalAvvPrimoTempo").value)
							If Val("" & Rec("GoalAvvSecondoTempo").value) <> -1 Then
								GoalSegnatiAvv += Val("" & Rec("GoalAvvSecondoTempo").value)
							End If
							If Val("" & Rec("GoalAvvTerzoTempo").value) <> -1 Then
								GoalSegnatiAvv += Val("" & Rec("GoalAvvTerzoTempo").value)
							End If

							Sql = "SELECT Count(*) FROM RisultatiAggiuntiviMarcatori Where idPartita=" & idPartita
							Rec2 = LeggeQuery(Conn, Sql, Connessione)
							Dim GoalSegnati As Integer = Rec2(0).Value
							Rec2.Close

							Dim g1 As Integer = Val(Mid(Risultato, 1, Risultato.IndexOf("-")))
							Dim g2 As Integer = Val(Mid(Risultato, Risultato.IndexOf("-") + 2, Risultato.Length))
							Dim gg1 As Integer = 0
							Dim gg2 As Integer = 0
							If RisGiochetti <> "" Then
								gg1 = Mid(RisGiochetti, 1, RisGiochetti.IndexOf("-"))
								gg2 = Mid(RisGiochetti, RisGiochetti.IndexOf("-") + 2, RisGiochetti.Length)

								g1 += gg1
								g2 += gg2
							End If
							g1 += GoalSegnati
							g2 += GoalSegnatiAvv

							If Rec("Casa").Value = "S" Then
								PartInCasa += 1
								GfInCasa += g1
								GSInCasa += g2
								If g1 > g2 Then
									VittInCasa += 1
								Else
									If g1 < g2 Then
										SconInCasa += 1
									Else
										PareInCasa += 1
									End If
								End If
								If Anno = idAnno Then
									PartInCasaAnno += 1
									GfInCasaAnno += g1
									GSInCasaAnno += g2
									If g1 > g2 Then
										VittInCasaAnno += 1
									Else
										If g1 < g2 Then
											SconInCasaAnno += 1
										Else
											PareInCasaAnno += 1
										End If
									End If
								End If
							Else
								PartFuoriCasa += 1
								GfFuoriCasa += g1
								GSFuoriCasa += g2
								If g1 > g2 Then
									VittFuoriCasa += 1
								Else
									If g1 < g2 Then
										SconFuoriCasa += 1
									Else
										PareFuoriCasa += 1
									End If
								End If
								If Anno = idAnno Then
									PartFuoriCasaAnno += 1
									GfFuoriCasaAnno += g1
									GSFuoriCasaAnno += g2
									If g1 > g2 Then
										VittFuoriCasaAnno += 1
									Else
										If g1 < g2 Then
											SconFuoriCasaAnno += 1
										Else
											PareFuoriCasaAnno += 1
										End If
									End If
								End If
							End If

							Rec.MoveNext
						Loop
					End If
					Rec.Close()

					Ritorno = IncontriTotali & ";" &
						PartInCasa & ";" & VittInCasa & ";" & PareInCasa & ";" & SconInCasa & ";" & GfInCasa & ";" & GSInCasa & ";" &
						PartFuoriCasa & ";" & VittFuoriCasa & ";" & PareFuoriCasa & ";" & SconFuoriCasa & ";" & GfFuoriCasa & ";" & GSFuoriCasa & ";" &
						IncontriAnno & ";" &
						PartInCasaAnno & ";" & VittInCasaAnno & ";" & PareInCasaAnno & ";" & SconInCasaAnno & ";" & GfInCasaAnno & ";" & GSInCasaAnno & ";" &
						PartFuoriCasaAnno & ";" & VittFuoriCasaAnno & ";" & PareFuoriCasaAnno & ";" & SconFuoriCasaAnno & ";" & GfFuoriCasaAnno & ";" & GSFuoriCasaAnno & ";"
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	Private Function EsegueStatistica(Conn As Object, Connessione As String, Sql As String) As Integer
        Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
        Dim Ritorno As Integer

        Rec = LeggeQuery(Conn, Sql, Connessione)
        If TypeOf (Rec) Is String Then
            Ritorno = -1
        Else
            If Rec.Eof Then
                Ritorno = -1
            Else
                Ritorno = Rec(0).Value
            End If
        End If
        Rec.Close()

        Return Ritorno
    End Function
End Class