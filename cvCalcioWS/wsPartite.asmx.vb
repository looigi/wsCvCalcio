Imports System.Web.Services
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_part.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsPartite
    Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function SalvaPartita(Squadra As String, idPartita As String, ByVal idAnno As String, ByVal idCategoria As String, ByVal idAvversario As String,
								 idAllenatore As String, DataOra As String, Casa As String, idTipologia As String,
								 idCampo As String, Risultato As String, Notelle As String, Marcatori As String, Convocati As String,
								 RisGiochetti As String, RisAvv As String, Campo As String, Tempo1Tempo As String,
								 Tempo2Tempo As String, Tempo3Tempo As String, Coordinate As String, sTempo As String,
								 idUnioneCalendario As String, TGA1 As String, TGA2 As String, TGA3 As String, Dirigenti As String, idArbitro As String,
								 RisultatoATempi As String, RigoriPropri As String, RigoriAvv As String, EventiPrimoTempo As String,
								 EventiSecondoTempo As String, EventiTerzoTempo As String) As String

		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Ok As Boolean = True
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Try
					Sql = "Delete From Partite Where idPartita=" & idPartita
					Ritorno = EsegueSql(Conn, Sql, Connessione)
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				If Ok Then
					Try
						Sql = "Delete From Risultati Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From Marcatori Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From Convocati Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From RisultatiAggiuntivi Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From RisultatiAggiuntiviMarcatori Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From CampiEsterni Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From CoordinatePartite Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From MeteoPartite Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From TempiGoalAvversari Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From DirigentiPartite Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From ArbitriPartite Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete * From RigoriAvversari Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete * From RigoriPropri Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete * From EventiPartita Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Dim d As Date = DataOra.Replace("%20", " ")
					d = d.AddHours(-1)
					Dim OraConv As String = Format(d.Hour, "00") & ":" & Format(d.Minute, "00") & ":" & Format(d.Second, "00")

					Try
						Sql = "Insert Into Partite Values (" &
							" " & idAnno & ", " &
							" " & idPartita & ", " &
							" " & idCategoria & ", " &
							" " & idAvversario & ", " &
							" " & idAllenatore & ", " &
							"'" & DataOra.Replace("%20", " ") & "', " &
							"'S', " &
							"'" & Casa & "', " &
							" " & idTipologia & ", " &
							" " & idCampo & ", " &
							"'" & OraConv & "', " &
							" " & idUnioneCalendario & ", " &
							"'" & RisultatoATempi & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					If Casa = "E" And Campo <> "" Then
						Try
							Sql = "Insert Into CampiEsterni Values (" &
							" " & idPartita & ", " &
							"'" & Campo.Replace("'", "''") & "' " &
							")"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					Try
						Sql = "Insert Into Risultati Values (" &
							" " & idPartita & ", " &
							"'" & Risultato & "', " &
							"'" & Notelle.Replace("'", "''") & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Dim GA() As String = RisAvv.Split(";")

					Try
						Sql = "Insert Into RisultatiAggiuntivi Values (" &
							" " & idPartita & ", " &
							"'" & RisGiochetti & "', " &
							" " & GA(0) & ", " &
							" " & GA(1) & ", " &
							" " & GA(2) & ", " &
							"'" & Tempo1Tempo & "', " &
							"'" & Tempo2Tempo & "', " &
							"'" & Tempo3Tempo & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Dim CC() As String = Coordinate.Split(";")

					Try
						Sql = "Insert Into CoordinatePartite Values (" &
							" " & idPartita & ", " &
							"'" & CC(0) & "', " &
							"'" & CC(1) & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Dim TT() As String = sTempo.Split(";")

					Try
						Sql = "Insert Into MeteoPartite Values (" &
							" " & idPartita & ", " &
							"'" & TT(0) & "', " &
							"'" & TT(1) & "', " &
							"'" & TT(2) & "', " &
							"'" & TT(3) & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Insert Into TempiGoalAvversari Values (" &
							" " & idPartita & ", " &
							"'" & TGA1 & "', " &
							"'" & TGA2 & "', " &
							"'" & TGA3 & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						For Each M As String In Marcatori.Split("§")
							If M <> "" Then
								Dim Campi() As String = M.Split(";")
								Dim Tempo As String = Campi(0)
								Dim idMarcatore As String = Campi(1)
								If Campi(3) = "Autorete" Then
									idMarcatore = -1
								End If
								Dim Minuto As String = ""
								If Campi.Length > 4 Then
									Minuto = Campi(5)
								End If

								If Minuto = "" Then Minuto = "null"

								Dim Progressivo As Integer = -1

								'Sql = "SELECT Max(idProgressivo)+1 FROM Marcatori Where idPartita=" & idPartita & " And idGiocatore=" & idMarcatore
								'Rec = LeggeQuery(Conn, Sql, Connessione)
								'If TypeOf (Rec) Is String Then
								'    Ritorno = Rec
								'    Ok = False
								'Else
								'    If Rec(0).Value Is DBNull.Value Then
								'        Progressivo = 1
								'    Else
								'        Progressivo = Rec(0).Value
								'    End If
								'    Rec.Close()
								'End If

								'If Ok Then
								'    Sql = "Insert Into Marcatori Values (" & _
								'        " " & idPartita & ", " &
								'        " " & idMarcatore & ", " &
								'        " " & Progressivo & ", " &
								'        " " & Minuto & " " &
								'        ")"
								'    Ritorno = EsegueSql(Conn, Sql, Connessione)
								'Else
								'    Exit For
								'End If

								If Ok Then
									Sql = "SELECT Max(Progressivo)+1 FROM RisultatiAggiuntiviMarcatori Where idPartita=" & idPartita & " And idTempo=" & Tempo
									Rec = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
									Else
										If Rec(0).Value Is DBNull.Value Then
											Progressivo = 1
										Else
											Progressivo = Rec(0).Value
										End If
										Rec.Close()
									End If

									Sql = "Insert Into RisultatiAggiuntiviMarcatori Values (" &
										" " & idPartita & ", " &
										" " & Tempo & ", " &
										" " & Progressivo & ", " &
										" " & idMarcatore & ", " &
										" " & Minuto & " " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
								End If
							End If
						Next
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Dim Progressivo As Integer = -1

					Sql = "SELECT Max(idProgressivo)+1 FROM Convocati Where idPartita=" & idPartita
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						If Rec(0).Value Is DBNull.Value Then
							Progressivo = 1
						Else
							Progressivo = Rec(0).Value
						End If
						Rec.Close()
					End If

					Try
						For Each C As String In Convocati.Split("§")
							If C <> "" Then
								Dim Campi() As String = C.Split(";")
								Dim idGioc As String = Campi(0)

								If Ok Then
									Sql = "Insert Into Convocati Values (" &
										" " & idPartita & ", " &
										" " & Progressivo & ", " &
										" " & idGioc & " " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)

									Progressivo += 1
								Else
									Exit For
								End If
							End If
						Next
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Dim Progressivo As Integer = -1

					Sql = "SELECT Max(Progressivo)+1 FROM DirigentiPartite Where idPartita=" & idPartita & " And idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						If Rec(0).Value Is DBNull.Value Then
							Progressivo = 1
						Else
							Progressivo = Rec(0).Value
						End If
						Rec.Close()
					End If

					Try
						For Each C As String In Dirigenti.Split("§")
							If C <> "" Then
								Dim Campi() As String = C.Split(";")
								Dim idDirigente As String = Campi(0)

								If Ok Then
									Sql = "Insert Into DirigentiPartite Values (" &
										" " & idAnno & ", " &
										" " & idPartita & ", " &
										" " & Progressivo & ", " &
										" " & idDirigente & " " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)

									Progressivo += 1
								Else
									Exit For
								End If
							End If
						Next
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						If idArbitro = 0 Then idArbitro = 1

						Sql = "Insert Into ArbitriPartite Values (" &
										" " & idAnno & ", " &
										" " & idPartita & ", " &
										"1, " &
										" " & idArbitro & " " &
										")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Dim RigPropri() As String = RigoriPropri.Split("§")
						Dim Conta As Integer = 0

						For Each s As String In RigPropri
							If s.Trim <> "" Then
								Dim c() As String = s.Split(";")

								Conta += 1
								Sql = "Insert Into RigoriPropri Values (" &
										" " & idAnno & ", " &
										" " & idPartita & ", " &
										" " & Conta & ", " &
										" " & c(0) & ", " &
										" " & c(6) & " " &
										")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
							End If
						Next
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						If RigoriAvv <> "" Then
							Dim a() As String = RigoriAvv.Split("§")

							Sql = "Insert Into RigoriAvversari Values (" &
									" " & idAnno & ", " &
									" " & idPartita & ", " &
									" " & a(0) & ", " &
									" " & a(1) & " " &
									")"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Dim e() As String = EventiPrimoTempo.Split("§")
						Dim progr As Integer = 0

						For Each ee As String In e
							If ee <> "" Then
								Dim eee() As String = ee.Split(";")
								Dim idEvento As String = eee(1)
								Dim idGiocatore As String = eee(3)
								Dim Minuto As String = eee(0)

								progr += 1
								Sql = "Insert Into EventiPartita Values (" &
									" " & idAnno & ", " &
									" " & idPartita & ", " &
									"1, " &
									" " & progr & ", " &
									" " & idEvento & ", " &
									" " & idGiocatore & ", " &
									" " & Minuto & " " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
							End If
						Next
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Dim e() As String = EventiSecondoTempo.Split("§")
						Dim progr As Integer = 0

						For Each ee As String In e
							If ee <> "" Then
								Dim eee() As String = ee.Split(";")
								Dim idEvento As String = eee(1)
								Dim idGiocatore As String = eee(3)
								Dim Minuto As String = eee(0)

								progr += 1
								Sql = "Insert Into EventiPartita Values (" &
										" " & idAnno & ", " &
										" " & idPartita & ", " &
										"2, " &
										" " & progr & ", " &
										" " & idEvento & ", " &
										" " & idGiocatore & ", " &
										" " & Minuto & " " &
										")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
							End If
						Next
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Dim e() As String = EventiTerzoTempo.Split("§")
						Dim progr As Integer = 0

						For Each ee As String In e
							If ee <> "" Then
								Dim eee() As String = ee.Split(";")
								Dim idEvento As String = eee(1)
								Dim idGiocatore As String = eee(3)
								Dim Minuto As String = eee(0)

								progr += 1
								Sql = "Insert Into EventiPartita Values (" &
										" " & idAnno & ", " &
										" " & idPartita & ", " &
										"3, " &
										" " & progr & ", " &
										" " & idEvento & ", " &
										" " & idGiocatore & ", " &
										" " & Minuto & " " &
										")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
							End If
						Next
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				CreaHtmlPartita(Squadra, Conn, Connessione, idAnno, idPartita)

				If Ok Then
					Ritorno = "*"
				End If
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaPartite(Squadra As String, idAnno As String, idCategoria As String) As String
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
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Try
					Sql = "SELECT Partite.DataOra, Partite.idPartita, Categorie.Descrizione As Categoria, SquadreAvversarie.Descrizione As Avversario, Risultati.Risultato, " &
						"Partite.Casa, Allenatori.Cognome+' '+Allenatori.Nome AS Allenatore, Partite.Casa As Casa, CampiAvversari.Descrizione+' '+CampiAvversari.Indirizzo As Campo, " &
						"Partite.idCategoria, Partite.idAvversario, Partite.idAllenatore, TipologiePartite.Descrizione As Tipologia, CampiEsterni.Descrizione As CampoEsterno, " &
						"AvversariCoord.Lat, AvversariCoord.Lon, Arbitri.idArbitro, Arbitri.Cognome +' '+Arbitri.Nome As Arbitro, Partite.RisultatoATempi " &
						"FROM ((((((((((Partite LEFT JOIN CampiAvversari ON Partite.idCampo = CampiAvversari.idCampo) " &
						"LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita) " &
						"LEFT JOIN Allenatori ON (Partite.idAnno = Allenatori.idAnno) AND (Partite.idAllenatore = Allenatori.idAllenatore)) " &
						"LEFT JOIN Categorie ON (Partite.idCategoria = Categorie.idCategoria) AND (Partite.idAnno = Categorie.idAnno)) " &
						"LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) " &
						"LEFT JOIN TipologiePartite ON Partite.idTipologia = TipologiePartite.idTipologia) " &
						"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita) " &
						"LEFT JOIN AvversariCoord ON Partite.idAvversario = AvversariCoord.idAvversario) " &
						"LEFT JOIN ArbitriPartite ON (Partite.idPartita = ArbitriPartite.idPartita And Partite.idAnno=ArbitriPartite.idAnno)) " &
						"LEFT JOIN Arbitri ON (Arbitri.idArbitro = ArbitriPartite.idArbitro And ArbitriPartite.idAnno = Arbitri.idAnno)) " &
						"WHERE Partite.idAnno=" & idAnno & " And Arbitri.idAnno=" & idAnno & " And Partite.Giocata='S' " &
						"And Partite.idCategoria=" & idCategoria & " Order By DataOra Desc"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna partita rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Dim Campo As String = Rec("Casa").Value

								If Campo = "S" Then
									Campo = "In casa"
								Else
									If Campo = "E" Then
										If Rec("CampoEsterno").Value Is DBNull.Value Then
											Campo = "Sconosciuto"
										Else
											Campo = Rec("CampoEsterno").Value
										End If
									Else
										If Rec("Campo").Value Is DBNull.Value Then
											Campo = "Sconosciuto"
										Else
											Campo = Rec("Campo").Value
										End If
									End If
								End If

								Ritorno &= Rec("DataOra").Value.ToString & ";"
								Ritorno &= Rec("idPartita").Value.ToString & ";"
								Ritorno &= Rec("Casa").Value.ToString & ";"
								Ritorno &= Rec("Categoria").Value.ToString & ";"
								If Rec("Avversario").Value Is DBNull.Value Then
									Ritorno &= "Sconosciuto" & ";"
								Else
									Ritorno &= Rec("Avversario").Value.ToString & ";"
								End If
								Ritorno &= Rec("Risultato").Value.ToString & ";"
								Ritorno &= Rec("Allenatore").Value.ToString & ";"
								Ritorno &= Campo & ";"
								Ritorno &= Rec("idCategoria").Value.ToString & ";"
								Ritorno &= Rec("idAvversario").Value.ToString & ";"
								Ritorno &= Rec("idAllenatore").Value & ";"
								Ritorno &= Rec("Tipologia").Value & ";"

								Dim goalAvversari As Integer = 0

								Sql = "Select GoalAvvPrimoTempo, GoalAvvSecondoTempo, GoalAvvTerzoTempo " &
									"From RisultatiAggiuntivi " &
									"Where idPartita=" & Rec("idPartita").Value.ToString
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									If Not Rec2.Eof Then
										If Rec2("GoalAvvPrimoTempo").Value > 0 Then
											goalAvversari += Rec2("GoalAvvPrimoTempo").Value
										End If
										If Rec2("GoalAvvSecondoTempo").Value > 0 Then
											goalAvversari += Rec2("GoalAvvSecondoTempo").Value
										End If
										If Rec2("GoalAvvTerzoTempo").Value > 0 Then
											goalAvversari += Rec2("GoalAvvTerzoTempo").Value
										End If
									End If
								End If
								Rec2.Close

								Dim goalPropri As Integer = 0

								Sql = "Select Count(*) As Goals " &
									"From RisultatiAggiuntiviMarcatori " &
									"Where idPartita=" & Rec("idPartita").Value.ToString
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									If Not Rec2.Eof Then
										If Not Rec2("Goals").Value Is DBNull.Value Then
											goalPropri = Rec2("Goals").Value
										End If
									End If
								End If
								Rec2.Close

								If Rec("Casa").Value.ToString.ToUpper = "S" Then
									Ritorno &= goalPropri.ToString.Trim & "-" & goalAvversari.ToString.Trim & ";"
								Else
									Ritorno &= goalAvversari.ToString.Trim & "-" & goalPropri.ToString.Trim & ";"
								End If

								Dim MultiMediaPartite As String = RitornaMultimediaPerTipologia(Squadra, idAnno, Rec("idPartita").Value, "Partite")

								If MultiMediaPartite <> "" Then
									Dim QuanteImmagini() As String = MultiMediaPartite.Split("§")
									Ritorno &= QuanteImmagini.Length.ToString & ";"
								Else
									Ritorno &= "0;"
								End If

								If Rec("Lat").Value.ToString <> "" And Rec("Lon").Value.ToString <> "" Then
									Ritorno &= Rec("Lat").Value.ToString & "," & Rec("Lon").Value.ToString & ";"
								Else
									Ritorno &= ";"
								End If

								Ritorno &= Rec("idArbitro").Value & ";"
								Ritorno &= Rec("Arbitro").Value & ";"
								Ritorno &= Rec("RisultatoATempi").Value & ";"

								Dim RigoriPropri As String = ""
								Dim RigoriAvv As String = "0!0!"

								Sql = "SELECT RigoriPropri.idGiocatore, Ruoli.Descrizione, Giocatori.Cognome + ' ' + Giocatori.Nome As Giocatore, " &
									"Giocatori.NumeroMaglia, RigoriPropri.Termine From ((RigoriPropri " &
									"Left Join Giocatori On RigoriPropri.idGiocatore=Giocatori.idGiocatore And RigoriPropri.idAnno = Giocatori.idAnno) " &
									"Left Join Ruoli On Giocatori.idRuolo = Ruoli.idRuolo) " &
									"Where RigoriPropri.idAnno=" & idAnno & " And idPartita=" & Rec("idPartita").Value.ToString & " " &
									"Order By RigoriPropri.idRigore"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									Do Until Rec2.Eof
										' 448;Centrocampista;Cataldi Lorenzo;;;14;-1;
										RigoriPropri &= Rec2("idGiocatore").Value & "!"
										RigoriPropri &= Rec2("Descrizione").Value & "!"
										RigoriPropri &= Rec2("Giocatore").Value & "!"
										RigoriPropri &= "!"
										RigoriPropri &= "!"
										RigoriPropri &= Rec2("NumeroMaglia").Value & "!"
										RigoriPropri &= Rec2("Termine").Value & "!"
										RigoriPropri &= "%"

										Rec2.MoveNext
									Loop
								End If
								Rec2.Close

								Sql = "Select * From RigoriAvversari Where idAnno=" & idAnno & " And idPartita=" & Rec("idPartita").Value.ToString
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									If Not Rec2.Eof Then
										RigoriAvv = Rec2("Segnati").Value & "!" & Rec2("Sbagliati").Value & "!"
									End If
								End If

								Ritorno &= RigoriPropri & ";"
								Ritorno &= RigoriAvv & ";"

								Ritorno &= "§"

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
	Public Function RitornaPartitaDaID(Squadra As String, idAnno As String, idPartita As String) As String
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
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Try
					Sql = "SELECT Partite.idPartita, Partite.idCategoria, Partite.idAvversario, Partite.idTipologia, Partite.idCampo, " &
						"Partite.idUnioneCalendario, Partite.DataOra, Partite.Giocata, Partite.OraConv, Risultati.Risultato, Risultati.Note, " &
						"RisultatiAggiuntivi.RisGiochetti, RisultatiAggiuntivi.GoalAvvPrimoTempo, RisultatiAggiuntivi.GoalAvvSecondoTempo, " &
						"RisultatiAggiuntivi.GoalAvvTerzoTempo, SquadreAvversarie.Descrizione AS Avversario, CampiAvversari.Descrizione AS Campo, " &
						"TipologiePartite.Descrizione AS Tipologia, Allenatori.Cognome+' '+Allenatori.Nome AS Allenatore, Categorie.Descrizione As Categoria, " &
						"CampiAvversari.Indirizzo as CampoIndirizzo, Partite.Casa, Allenatori.idAllenatore, CampiEsterni.Descrizione As CampoEsterno, " &
						"RisultatiAggiuntivi.Tempo1Tempo, RisultatiAggiuntivi.Tempo2Tempo, RisultatiAggiuntivi.Tempo3Tempo, " &
						"CoordinatePartite.Lat, CoordinatePartite.Lon, TempiGoalAvversari.TempiPrimoTempo, TempiGoalAvversari.TempiSecondoTempo, TempiGoalAvversari.TempiTerzoTempo, " &
						"MeteoPartite.Tempo, MeteoPartite.Gradi, MeteoPartite.Umidita, MeteoPartite.Pressione, ArbitriPartite.idArbitro, Arbitri.Cognome + ' ' + Arbitri.Nome As Arbitro, " &
						"Partite.RisultatoATempi " &
						"FROM ((((((((((((Partite LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita) " &
						"LEFT JOIN RisultatiAggiuntivi ON Partite.idPartita = RisultatiAggiuntivi.idPartita) " &
						"LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) " &
						"LEFT JOIN TipologiePartite ON Partite.idTipologia = TipologiePartite.idTipologia) " &
						"LEFT JOIN Allenatori ON (Partite.idAnno = Allenatori.idAnno) And (Partite.idAllenatore = Allenatori.idAllenatore)) " &
						"LEFT JOIN CampiAvversari ON SquadreAvversarie.idCampo = CampiAvversari.idCampo) " &
						"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita) " &
						"LEFT JOIN Categorie ON Partite.idCategoria = Categorie.idCategoria And Categorie.idAnno = Partite.idAnno) " &
						"LEFT JOIN CoordinatePartite On Partite.idPartita = CoordinatePartite.idPartita) " &
						"LEFT JOIN MeteoPartite On Partite.idPartita = MeteoPartite.idPartita) " &
						"LEFT JOIN TempiGoalAvversari On Partite.idPartita = TempiGoalAvversari.idPartita) " &
						"LEFT JOIN ArbitriPartite On (Partite.idPartita = ArbitriPartite.idPartita And ArbitriPartite.idAnno = Partite.idAnno)) " &
						"LEFT JOIN Arbitri On ArbitriPartite.idArbitro=Arbitri.idArbitro And ArbitriPartite.idAnno=Arbitri.idAnno " &
						"WHERE Partite.idPartita=" & idPartita & " And Partite.idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Sql & "--->" & Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " No partites found"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idCategoria").Value.ToString & ";" &
									Rec("idAvversario").Value.ToString & ";" &
									Rec("idTipologia").Value.ToString & ";" &
									Rec("idCampo").Value.ToString & ";" &
									Rec("idUnioneCalendario").Value.ToString & ";" &
									Rec("DataOra").Value.ToString & ";" &
									Rec("Giocata").Value.ToString & ";" &
									Rec("OraConv").Value.ToString & ";" &
									Rec("Note").Value.ToString & ";" &
									Rec("RisGiochetti").Value.ToString & ";" &
									Rec("GoalAvvPrimoTempo").Value.ToString & ";" &
									Rec("GoalAvvSecondoTempo").Value.ToString & ";" &
									Rec("GoalAvvTerzoTempo").Value.ToString & ";" &
									Rec("Avversario").Value.ToString & ";"
								If Rec("Casa").Value = "E" Then
									Ritorno &= Rec("CampoEsterno").Value.ToString & ";"
								Else
									Ritorno &= Rec("Campo").Value.ToString & ";"
								End If
								Ritorno &= Rec("Allenatore").Value.ToString & ";" &
									Rec("Categoria").Value.ToString & ";" &
									Rec("CampoIndirizzo").Value.ToString & ";" &
									Rec("Tipologia").Value.ToString & ";" &
									Rec("Casa").Value.ToString & ";" &
									Rec("idAllenatore").Value.ToString & ";"

								Dim goalAvversari As Integer = 0

								Sql = "Select GoalAvvPrimoTempo, GoalAvvSecondoTempo, GoalAvvTerzoTempo " &
									"From RisultatiAggiuntivi " &
									"Where idPartita=" & Rec("idPartita").Value.ToString
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									If Not Rec2.Eof Then
										If Rec2("GoalAvvPrimoTempo").Value > 0 Then
											goalAvversari += Rec2("GoalAvvPrimoTempo").Value
										End If
										If Rec2("GoalAvvSecondoTempo").Value > 0 Then
											goalAvversari += Rec2("GoalAvvSecondoTempo").Value
										End If
										If Rec2("GoalAvvTerzoTempo").Value > 0 Then
											goalAvversari += Rec2("GoalAvvTerzoTempo").Value
										End If
									End If
								End If
								Rec2.Close

								Dim Dirigenti As String = ""

								Sql = "Select Dirigenti.idDirigente, Dirigenti.Cognome + ' ' + Dirigenti.Nome As Dirigente " &
									"From DirigentiPartite " &
									"Left Join Dirigenti On (DirigentiPartite.idAnno=Dirigenti.idAnno And DirigentiPartite.idDirigente=Dirigenti.idDirigente) " &
									"Where DirigentiPartite.idPartita=" & Rec("idPartita").Value.ToString & " And DirigentiPartite.idAnno=" & idAnno
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									Do Until Rec2.Eof
										Dirigenti &= Rec2("idDirigente").Value & "!" & Rec2("Dirigente").Value & "%"

										Rec2.MoveNext()
									Loop
									Rec2.Close
								End If

								Dim goalPropri As Integer = 0

								Sql = "Select Count(*) As Goals " &
									"From RisultatiAggiuntiviMarcatori " &
									"Where idPartita=" & Rec("idPartita").Value.ToString
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									If Not Rec2.Eof Then
										If Not Rec2("Goals").Value Is DBNull.Value Then
											goalPropri = Rec2("Goals").Value
										End If
									End If
								End If
								Rec2.Close

								If Rec("Casa").Value.ToString.ToUpper = "S" Then
									Ritorno &= goalPropri.ToString.Trim & "-" & goalAvversari.ToString.Trim & ";"
								Else
									Ritorno &= goalAvversari.ToString.Trim & "-" & goalPropri.ToString.Trim & ";"
								End If

								Ritorno &= Rec("Tempo1Tempo").Value & ";"
								Ritorno &= Rec("Tempo2Tempo").Value & ";"
								Ritorno &= Rec("Tempo3Tempo").Value & ";"

								Ritorno &= Rec("Lat").Value & ";"
								Ritorno &= Rec("Lon").Value & ";"

								Ritorno &= Rec("Tempo").Value & ";"
								Ritorno &= Rec("Gradi").Value & ";"
								Ritorno &= Rec("Umidita").Value & ";"
								Ritorno &= Rec("Pressione").Value & ";"

								Ritorno &= Rec("TempiPrimoTempo").Value & ";"
								Ritorno &= Rec("TempiSecondoTempo").Value & ";"
								Ritorno &= Rec("TempiTerzoTempo").Value & ";"

								Ritorno &= Dirigenti & ";"

								Ritorno &= Rec("idArbitro").Value.ToString & "-" & Rec("Arbitro").Value.ToString & ";"

								Ritorno &= Rec("RisultatoATempi").Value.ToString & ";"

								Dim RigoriPropri As String = ""
								Dim RigoriAvv As String = "0!0!"

								Sql = "SELECT RigoriPropri.idGiocatore, Ruoli.Descrizione, Giocatori.Cognome + ' ' + Giocatori.Nome As Giocatore, " &
									"Giocatori.NumeroMaglia, RigoriPropri.Termine From ((RigoriPropri " &
									"Left Join Giocatori On RigoriPropri.idGiocatore=Giocatori.idGiocatore And RigoriPropri.idAnno = Giocatori.idAnno) " &
									"Left Join Ruoli On Giocatori.idRuolo = Ruoli.idRuolo) " &
									"Where RigoriPropri.idAnno=" & idAnno & " And idPartita=" & Rec("idPartita").Value.ToString & " " &
									"Order By RigoriPropri.idRigore"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									Do Until Rec2.Eof
										' 448;Centrocampista;Cataldi Lorenzo;;;14;-1;
										RigoriPropri &= Rec2("idGiocatore").Value & "!"
										RigoriPropri &= Rec2("Descrizione").Value & "!"
										RigoriPropri &= Rec2("Giocatore").Value & "!"
										RigoriPropri &= "!"
										RigoriPropri &= "!"
										RigoriPropri &= Rec2("NumeroMaglia").Value & "!"
										RigoriPropri &= Rec2("Termine").Value & "!"
										RigoriPropri &= "%"

										Rec2.MoveNext
									Loop
								End If
								Rec2.Close

								Sql = "Select * From RigoriAvversari Where idAnno=" & idAnno & " And idPartita=" & Rec("idPartita").Value.ToString
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									If Not Rec2.Eof Then
										RigoriAvv = Rec2("Segnati").Value & "!" & Rec2("Sbagliati").Value & "!"
									End If
								End If

								Ritorno &= RigoriPropri & ";"
								Ritorno &= RigoriAvv & ";"

								Dim EventiPrimoTempo As String = ""

								Sql = "Select EventiPartita.Minuto, EventiPartita.idEvento, Eventi.Descrizione, EventiPartita.idGiocatore, iif(Giocatori.Cognome + ' ' + Giocatori.Nome is null, 'Avversario', Giocatori.Cognome + ' ' + Giocatori.Nome) As Giocatore " &
									"From ((EventiPartita LEFT Join Eventi On Eventi.idEvento=EventiPartita.idEvento) " &
									"LEFT JOIN Giocatori On EventiPartita.idAnno=Giocatori.idAnno And EventiPartita.idGiocatore=Giocatori.idGiocatore) " &
									"Where EventiPartita.idAnno=" & idAnno & " And idPartita=" & Rec("idPartita").Value.ToString & " And idTempo=1 Order By Progressivo"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									Do Until Rec2.Eof
										EventiPrimoTempo &= Rec2("Minuto").Value & "!" &
											Rec2("idEvento").Value & "!" &
											Rec2("Descrizione").Value & "!" &
											Rec2("idGiocatore").Value & "!" &
											Rec2("Giocatore").Value & "!%"

										Rec2.MoveNext
									Loop
									Rec2.Close
								End If

								Dim EventiSecondoTempo As String = ""

								Sql = "Select EventiPartita.Minuto, EventiPartita.idEvento, Eventi.Descrizione, EventiPartita.idGiocatore, iif(Giocatori.Cognome + ' ' + Giocatori.Nome is null, 'Avversario', Giocatori.Cognome + ' ' + Giocatori.Nome) As Giocatore " &
									"From ((EventiPartita LEFT Join Eventi On Eventi.idEvento=EventiPartita.idEvento) " &
									"LEFT JOIN Giocatori On EventiPartita.idAnno=Giocatori.idAnno And EventiPartita.idGiocatore=Giocatori.idGiocatore) " &
									"Where EventiPartita.idAnno=" & idAnno & " And idPartita=" & Rec("idPartita").Value.ToString & " And idTempo=2 Order By Progressivo"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									Do Until Rec2.Eof
										EventiSecondoTempo &= Rec2("Minuto").Value & "!" &
											Rec2("idEvento").Value & "!" &
											Rec2("Descrizione").Value & "!" &
											Rec2("idGiocatore").Value & "!" &
											Rec2("Giocatore").Value & "!%"

										Rec2.MoveNext
									Loop
									Rec2.Close
								End If

								Dim EventiTerzoTempo As String = ""

								Sql = "Select EventiPartita.Minuto, EventiPartita.idEvento, Eventi.Descrizione, EventiPartita.idGiocatore, iif(Giocatori.Cognome + ' ' + Giocatori.Nome is null, 'Avversario', Giocatori.Cognome + ' ' + Giocatori.Nome) As Giocatore " &
									"From ((EventiPartita LEFT Join Eventi On Eventi.idEvento=EventiPartita.idEvento) " &
									"LEFT JOIN Giocatori On EventiPartita.idAnno=Giocatori.idAnno And EventiPartita.idGiocatore=Giocatori.idGiocatore) " &
									"Where EventiPartita.idAnno=" & idAnno & " And idPartita=" & Rec("idPartita").Value.ToString & " And idTempo=3 Order By Progressivo"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									Do Until Rec2.Eof
										EventiTerzoTempo &= Rec2("Minuto").Value & "!" &
											Rec2("idEvento").Value & "!" &
											Rec2("Descrizione").Value & "!" &
											Rec2("idGiocatore").Value & "!" &
											Rec2("Giocatore").Value & "!%"

										Rec2.MoveNext
									Loop
									Rec2.Close
								End If

								Ritorno &= EventiPrimoTempo & ";"
								Ritorno &= EventiSecondoTempo & ";"
								Ritorno &= EventiTerzoTempo & ";"

								Ritorno &= "§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()
						Ritorno &= "|"

						Sql = "Select * From (Select idTempo, Progressivo, RisultatiAggiuntiviMarcatori.idGiocatore, Minuto, Cognome, Nome, Ruoli.Descrizione As Ruolo, NumeroMaglia " &
							"FROM ((RisultatiAggiuntiviMarcatori " &
							"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore) " &
							"Left Join Ruoli On Giocatori.idRuolo = Ruoli.idRuolo) " &
							"Where RisultatiAggiuntiviMarcatori.idPartita=" & idPartita & " And Giocatori.idAnno=" & idAnno & " " &
							"Union All " &
							"Select idTempo, Progressivo, -1, Minuto, 'Autorete' As Cognome, '' As Nome, '' As Ruolo, 999 As NumeroMaglia FROM RisultatiAggiuntiviMarcatori " &
							"Where RisultatiAggiuntiviMarcatori.idPartita = " & idPartita & " And RisultatiAggiuntiviMarcatori.idGiocatore = -1 " &
							") As A  Order By idTempo, Progressivo"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Sql & "--->" & Rec
						Else
							If Rec.Eof Then
								'Ritorno &= "|"
							Else
								Do Until Rec.Eof
									Ritorno &= Rec("idTempo").Value.ToString & ";" &
										Rec("Progressivo").Value.ToString & ";" &
										Rec("idGiocatore").Value.ToString & ";" &
										Rec("Minuto").Value.ToString & ";" &
										Rec("Cognome").Value.ToString & ";" &
										Rec("Nome").Value.ToString & ";" &
										Rec("Ruolo").Value.ToString & ";" &
										Rec("NumeroMaglia").Value.ToString & ";" &
										"§"

									Rec.MoveNext()
								Loop
								'Ritorno &= "|"
							End If
							Rec.Close()
						End If

						Sql = "SELECT idProgressivo, Marcatori.idGiocatore, Minuto, Cognome, Nome, Ruoli.Descrizione As Ruolo, NumeroMaglia " &
							"FROM ((Marcatori " &
							"Left Join Giocatori On Marcatori.idGiocatore = Giocatori.idGiocatore) " &
							"Left Join Ruoli On Giocatori.idRuolo = Ruoli.idRuolo) " &
							"Where Marcatori.idPartita=" & idPartita & " And Giocatori.idAnno=" & idAnno & " Order By idProgressivo"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Sql & "--->" & Rec
						Else
							If Rec.Eof Then
								Ritorno &= "|"
							Else
								Do Until Rec.Eof
									Ritorno &= "1;" &
										Rec("idProgressivo").Value.ToString & ";" &
										Rec("idGiocatore").Value.ToString & ";" &
										Rec("Minuto").Value.ToString & ";" &
										Rec("Cognome").Value.ToString & ";" &
										Rec("Nome").Value.ToString & ";" &
										Rec("Ruolo").Value.ToString & ";" &
										Rec("NumeroMaglia").Value.ToString & ";" &
										"§"

									Rec.MoveNext()
								Loop
								Ritorno &= "|"
							End If
							Rec.Close()
						End If

						Sql = "SELECT idProgressivo, Convocati.idGiocatore, Cognome, Nome, Ruoli.idRuolo, Ruoli.Descrizione As Ruolo, NumeroMaglia " &
							"FROM ((Convocati " &
							"Left Join Giocatori On Convocati.idGiocatore = Giocatori.idGiocatore) " &
							"Left Join Ruoli On Giocatori.idRuolo = Ruoli.idRuolo) " &
							"Where Convocati.idPartita=" & idPartita & " And Giocatori.idAnno=" & idAnno & " Order By idProgressivo"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Sql & "--->" & Rec
						Else
							If Rec.Eof Then
								Ritorno &= "|"
							Else
								Do Until Rec.Eof
									Ritorno &= Rec("idProgressivo").Value.ToString & ";" &
										Rec("idGiocatore").Value.ToString & ";" &
										Rec("Cognome").Value.ToString & ";" &
										Rec("Nome").Value.ToString & ";" &
										Rec("Ruolo").Value.ToString & ";" &
										Rec("idRuolo").Value.ToString & ";" &
										Rec("NumeroMaglia").Value.ToString & ";" &
										"§"

									Rec.MoveNext()
								Loop
								Ritorno &= "|"
							End If
							Rec.Close()
						End If
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
	Public Function RitornaIdPartita(Squadra As String) As String
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
				Dim idPartita As Integer

				Try
					Sql = "SELECT Max(idPartita)+1 FROM Partite"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec(0).Value Is DBNull.Value Then
							idPartita = 1
						Else
							idPartita = Rec(0).Value
						End If
						Rec.Close()
					End If
					Ritorno = idPartita.ToString
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function


	<WebMethod()>
	Public Function CreaFoglioConvocazioni(Squadra As String, idAnno As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim gf As New GestioneFilesDirectory

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Sql = "SELECT Partite.idPartita, Partite.DataOra, Categorie.Descrizione As Categoria, SquadreAvversarie.Descrizione As Avversario, '' + CampiAvversari.Descrizione As Campo, " &
					"CampiAvversari.Indirizzo, '' + CampiEsterni.Descrizione As CampoEsterno, Allenatori.Cognome + ' ' + Allenatori.Nome As Mister, Allenatori.Telefono, Partite.OraConv, " &
					"Anni.Descrizione As Anno, Anni.Indirizzo As IndCasa, Anni.CampoSquadra, Partite.Casa " &
					"FROM (((((Partite LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) " &
					"LEFT JOIN CampiAvversari ON SquadreAvversarie.idCampo = CampiAvversari.idCampo) " &
					"LEFT JOIN Categorie ON (Partite.idAnno = Categorie.idAnno) And (Partite.idCategoria = Categorie.idCategoria)) " &
					"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita) " &
					"LEFT JOIN Allenatori ON (Partite.idAllenatore = Allenatori.idAllenatore) And (Partite.idCategoria = Allenatori.idCategoria) And (Partite.idAnno = Allenatori.idAnno)) " &
					"LEFT JOIN Anni On (Anni.idAnno = Partite.idAnno) " &
					"WHERE Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = "ERROR: Nessuna partita rilevata"
					Else
						Dim Filetto As String = gf.LeggeFileIntero(Server.MapPath(".") & "\base_convocazioni.txt")
						Dim Datella As Date = "" & Rec("DataOra").Value
						Dim DatellaConv As Date = Datella.AddMinutes(-75)

						Filetto = Filetto.Replace("***PARTITA***", "Partita " & idPartita & " Anno " & Rec("Anno").Value)
						Filetto = Filetto.Replace("***NOME_SQUADRA***", Squadra.ToUpper.Trim.Replace("_", " "))
						Filetto = Filetto.Replace("***LOGO***", "<img src=""http://looigi.no-ip.biz:12345/wsCvCalcio/Loghi/" & Squadra & ".png"" style=""width: 80px; height: 100px;"" />")
						Filetto = Filetto.Replace("***SQUADRA***", "" & Rec("Categoria").Value)
						Filetto = Filetto.Replace("***GARA***", "" & Rec("Categoria").Value & " - " & "" & Rec("Avversario").Value)
						Filetto = Filetto.Replace("***DATA***", Format(Datella.Day, "00") & "/" & Format(Datella.Month, "00") & "/" & Datella.Year)
						If "" & Rec("Casa").Value = "E" Then
							Filetto = Filetto.Replace("***CAMPO***", "" & Rec("CampoEsterno").Value)
							Filetto = Filetto.Replace("***INDIRIZZO***", "")
						Else
							If "" & Rec("Casa").Value = "S" Then
								Filetto = Filetto.Replace("***CAMPO***", "" & Rec("CampoSquadra").Value)
								Filetto = Filetto.Replace("***INDIRIZZO***", "" & Rec("IndCasa").Value)
							Else
								Filetto = Filetto.Replace("***CAMPO***", "" & Rec("Campo").Value)
								Filetto = Filetto.Replace("***INDIRIZZO***", "" & Rec("Indirizzo").Value)
							End If
						End If
						Filetto = Filetto.Replace("***ORARIO1***", Format(DatellaConv.Hour, "00") & ":" & Format(DatellaConv.Minute, "00"))
						Filetto = Filetto.Replace("***ORARIO2***", Format(Datella.Hour, "00") & ":" & Format(Datella.Minute, "00"))

						Filetto = Filetto.Replace("***MISTER***", "" & Rec("Mister").Value)
						Filetto = Filetto.Replace("***CELL***", "" & Rec("Telefono").Value)

						Rec.Close()

						Dim Convocati As String = ""

						Sql = "SELECT Giocatori.Cognome+' '+Giocatori.Nome AS Giocatore, Ruoli.idRuolo " &
							"FROM (Convocati LEFT JOIN Giocatori ON Convocati.idGiocatore = Giocatori.idGiocatore) LEFT JOIN Ruoli ON Giocatori.idRuolo = Ruoli.idRuolo " &
							"WHERE Convocati.idPartita=" & idPartita & " AND Giocatori.idAnno=" & idAnno & " " &
							"ORDER BY Ruoli.idRuolo, Giocatori.Cognome, Giocatori.Nome"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = "ERROR: Nessuna partita rilevata"
							Else
								Dim Giocatori As List(Of String) = New List(Of String)

								Do Until Rec.Eof
									Giocatori.Add("" & Rec("Giocatore").Value)

									Rec.MoveNext
								Loop
								Rec.Close

								Convocati &= "<table style=""width: 100%;"" cellpadding=""0px"" cellspacing=""0px"">"
								For i As Integer = 1 To 12
									Dim Riga As String = "<tr>"

									Riga &= "<td colspan=""1"" class=""adestra"">"
									Riga &= "<span class=""titolo3"">" & i & "</span>"
									Riga &= "</td>"

									Riga &= "<td style=""width:10px;"">"
									Riga &= "&nbsp;"
									Riga &= "</td>"

									Riga &= "<td colspan=""4"">"
									If i - 1 < Giocatori.Count Then
										Riga &= "<span class=""titolo3"">" & Giocatori.Item(i - 1) & "</span>"
									Else
										Riga &= "<div style=""width: 100""><span class=""titolo3"">&nbsp;</span></div>"
									End If
									Riga &= "</td>"

									Riga &= "<td colspan=""1"" class=""adestra"">"
									Riga &= "<span class=""titolo3"">" & i + 12 & "</span>"
									Riga &= "</td>"

									Riga &= "<td style=""width:10px;"">"
									Riga &= "&nbsp;"
									Riga &= "</td>"

									Riga &= "<td colspan=""4"">"
									If i + 11 < Giocatori.Count Then
										Riga &= "<span class=""titolo3"">" & Giocatori.Item(i + 11) & "</span>"
									Else
										Riga &= "<div style=""width: 100""><span class=""titolo3"">&nbsp;</span></div>"
									End If
									Riga &= "</td>"
									Riga &= "</tr>"

									Convocati &= Riga
								Next
								Convocati &= "</table>"
							End If
							Filetto = Filetto.Replace("***CONVOCATI***", Convocati)

							gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Convocazioni\" & Squadra & "\")
							gf.CreaAggiornaFile(Server.MapPath(".") & "\Convocazioni\" & Squadra & "\" & idPartita & ".html", Filetto)

							Ritorno = "*"
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class