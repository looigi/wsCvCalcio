Imports System.Web.Services
Imports System.ComponentModel
Imports System.Net.Mail
Imports System.Data.OleDb
Imports System.Web.ApplicationServices
Imports System.Web.Hosting
Imports System.IO
Imports System.Security.Principal

<System.Web.Services.WebService(Namespace:="http://cvcalcio_part.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
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
								 EventiSecondoTempo As String, EventiTerzoTempo As String, Mittente As String, DataOraAppuntamento As String, LuogoAppuntamento As String,
								 MezzoTrasporto As String, MandaMail As String, InFormazione As String, ShootOut As String, Tempi As String, PartitaConRigori As String,
								 idCapitano As String, CreaSchedaPartita As String, TempiGoalAvversari1T As String, TempiGoalAvversari2T As String, TempiGoalAvversari3T As String) As String

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

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
				If Ritorno <> "*" Then
					Ok = False
				End If

				If Ok Then
					Try
						Sql = "Delete From RisultatiAvversariMinuti Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From Partite Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From Risultati Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From Marcatori Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From Convocati Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From RisultatiAggiuntivi Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From RisultatiAggiuntiviMarcatori Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From CampiEsterni Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From CoordinatePartite Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From MeteoPartite Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From TempiGoalAvversari Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From DirigentiPartite Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From ArbitriPartite Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "delete from RigoriAvversari Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete from RigoriPropri Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete from EventiPartita Where idPartita=" & idPartita & " And idAnno=" & idAnno
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From InFormazione Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Try
						Sql = "Delete From PartiteCapitani Where idPartita=" & idPartita
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

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
							"'" & DataOra.Replace("%20", " ").Replace("-", "") & "', " &
							"'S', " &
							"'" & Casa & "', " &
							" " & idTipologia & ", " &
							" " & idCampo & ", " &
							"'" & OraConv & "', " &
							" " & idUnioneCalendario & ", " &
							"'" & RisultatoATempi & "', " &
							"'" & DataOraAppuntamento & "', " &
							"'" & LuogoAppuntamento.Replace("-", "") & "', " &
							"'" & MezzoTrasporto & "', " &
							"'" & ShootOut & "',  " &
							" " & Tempi & ", " &
							"'" & PartitaConRigori & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
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
					If idCapitano <> "" And idCapitano <> "-1" Then
						Try
							Sql = "Insert Into PartiteCapitani Values (" &
						" " & idPartita & ", " &
						" " & idCapitano & " " &
						")"
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

				If Ok Then
					Try
						Sql = "Insert Into Risultati Values (" &
							" " & idPartita & ", " &
							"'" & Risultato & "', " &
							"'" & Notelle.Replace("'", "''") & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					If RisAvv <> "" Then
						Dim GA() As String = RisAvv.Split(";")

						Try
							Sql = "Insert Into RisultatiAggiuntivi Values (" &
								" " & idPartita & ", " &
								"'" & RisGiochetti & "', " &
								" " & IIf(GA(0).Trim <> "", GA(0), "0") & ", " &
								" " & IIf(GA(1).Trim <> "", GA(1), "0") & ", " &
								" " & IIf(GA(2).Trim <> "", GA(2), "0") & ", " &
								"'" & Tempo1Tempo & "', " &
								"'" & Tempo2Tempo & "', " &
								"'" & Tempo3Tempo & "' " &
								")"
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

				If Ok Then
					If Coordinate <> "" Then
						Dim CC() As String = Coordinate.Split(";")

						Try
							Sql = "Insert Into CoordinatePartite Values (" &
								" " & idPartita & ", " &
								"'" & CC(0) & "', " &
								"'" & CC(1) & "' " &
								")"
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

				If Ok Then
					If sTempo <> "" Then
						If Coordinate <> "" Then
							Dim CC() As String = Coordinate.Split(";")
							Dim TempoMeteo As String = RitornaMeteo(CC(0), CC(1))

							If TempoMeteo.Contains(StringaErrore) Then
								Ritorno = TempoMeteo
								TempoMeteo = ""
								'Ok = False
							End If

							If TempoMeteo <> "" Then
								Dim TT() As String = TempoMeteo.Split(";")

								'Temperatura
								'Umidita
								'Pressione
								'Tempo
								'Icona

								Try
									Sql = "Insert Into MeteoPartite Values (" &
										" " & idPartita & ", " &
										"'" & TT(0) & "', " &
										"'" & TT(1) & "', " &
										"'" & TT(2) & "', " &
										"'" & TT(3) & "', " &
										"'" & TT(4) & "' " &
										")"
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
				End If

				If Ok Then
					Try
						Sql = "Insert Into TempiGoalAvversari Values (" &
							" " & idPartita & ", " &
							"'" & TGA1.Replace("$", "#") & "', " &
							"'" & TGA2.Replace("$", "#") & "', " &
							"'" & TGA3.Replace("$", "#") & "' " &
							")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					If Marcatori <> "" Then
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
									Dim Rigore As String = Campi(6)

									If Minuto = "" Then Minuto = "null"

									Dim Progressivo As Integer = -1

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
											" " & Minuto & ", " &
											"'" & Rigore & "' " &
											")"
										Ritorno = EsegueSql(Conn, Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
											Exit For
										End If
									End If
								End If
							Next
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					If Convocati <> "" Then
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

						' Prende mails dei convocati
						Dim convString As String = ""
						Dim MailsConvocati As New List(Of String)

						For Each C As String In Convocati.Split("§")
							If C <> "" Then
								Dim Campi() As String = C.Split(";")
								Dim idGioc As String = Campi(0)

								convString &= idGioc & ","
							End If
						Next
						If convString <> "" Then
							convString = Mid(convString, 1, convString.Length - 1)
						End If

						'Sql = "Select A.idGiocatore, Cognome, Nome, EMail From Giocatori A " &
						'	"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						'	"Where A.idGiocatore in (" & convString & ")"
						Sql = "Select A.Mail, A.idGiocatore, A.Progressivo, B.Cognome, B.Nome From GiocatoriMails A " &
							"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
							"Where A.idGiocatore in (" & convString & ") And Attiva = 'S' And A.Mail <> ''"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
						Else
							Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")

							Do Until Rec.Eof()
								If "" & Rec("Mail").Value <> "" Then
									Sql = "Select * From GiocatoriDettaglio A " &
										"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
										"Where A.idGiocatore=" & Rec("idGiocatore").Value & " " &
										"Order By Cognome, Nome"
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If Not Rec2.Eof Then
										Dim genitore As String = ""
										Select Case Val("" & Rec("Progressivo").Value)
											Case 1
												genitore = "" & Rec2("Genitore1").Value
											Case 2
												genitore = "" & Rec2("Genitore2").Value
											Case 3
												genitore = "" & Rec("Cognome").Value & " " & Rec("Nome").Value
										End Select
										'Dim cognome As String = ""
										'Dim nome As String = ""
										'If genitore.Contains(" ") Then
										'	Dim g() As String = genitore.Split(" ")
										'	cognome = g(0)
										'	nome = g(1)
										'Else
										'	cognome = genitore
										'	nome = ""
										'End If

										MailsConvocati.Add(Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("Mail").Value & ";C;" & Rec("idGiocatore").Value)
									End If
									Rec2.Close
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()
						End If

						If Ok Then
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
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
												Exit For
											End If

											Progressivo += 1
										Else
											Exit For
										End If
									End If
								Next

								Dim ma As New mail
								Dim Avversario As String = ""
								Dim Lat As String = ""
								Dim Lon As String = ""
								Dim Telefono As String = ""
								Dim Referente As String = ""
								Dim Categoria As String = ""
								Dim Anticipo As Single
								Dim sCampo As String = ""
								Dim IndirizzoCampo As String = ""
								Dim Allenatore As String = ""
								Dim TelAllenatore As String = ""
								Dim tipoPartita As String = ""

								Sql = "Select * From SquadreAvversarie A Left Join AvversariCoord B On A.idAvversario = B.idAvversario Where A.idAvversario = " & idAvversario
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
									Ok = False
								Else
									If Not Rec.Eof() Then
										Avversario = "" & Rec("Descrizione").Value
										Telefono = "" & Rec("Telefono").Value
										Referente = "" & Rec("Referente").Value
										Lat = IIf(Rec("Lat").Value Is DBNull.Value, 0, Rec("Lat").Value)
										Lon = IIf(Rec("Lon").Value Is DBNull.Value, 0, Rec("Lon").Value)
									End If
									Rec.Close()
								End If

								Sql = "Select * From Categorie Where idCategoria = " & idCategoria
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
									Ok = False
								Else
									If Not Rec.Eof() Then
										Categoria = "" & Rec("Descrizione").Value
										Anticipo = Rec("AnticipoConvocazione").Value
									End If
									Rec.Close()
								End If

								Sql = "Select * From [Generale].[dbo].[TipologiePartite] Where idTipologia = " & idTipologia
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
									Ok = False
								Else
									If Not Rec.Eof() Then
										tipoPartita = Rec("Descrizione").Value
									End If
									Rec.Close()
								End If

								Sql = "Select * From Allenatori Where idAllenatore = " & idAllenatore
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
									Ok = False
								Else
									If Not Rec.Eof() Then
										Allenatore = Rec("Cognome").Value & " " & Rec("Nome").Value
										TelAllenatore = "" & Rec("Telefono").Value
										MailsConvocati.Add(Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("EMail").Value & ";A;" & idAllenatore)
									End If
									Rec.Close()
								End If

								If Casa = "S" Then
									Sql = "Select * From Anni Where idAnno = " & idAnno
									Rec = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
									Else
										If Not Rec.Eof() Then
											sCampo = "" & Rec("CampoSquadra").Value
											IndirizzoCampo = "" & Rec("Indirizzo").Value
											Telefono = "" & Rec("Telefono").Value
											Lat = IIf(Rec("Lat").Value Is DBNull.Value, 0, Rec("Lat").Value)
											Lon = IIf(Rec("Lon").Value Is DBNull.Value, 0, Rec("Lon").Value)
											Referente = ""
										End If
										Rec.Close()
									End If
								Else
									Sql = "Select * From CampiAvversari Where idCampo = " & idCampo
									Rec = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
									Else
										If Not Rec.Eof() Then
											sCampo = "" & Rec("Descrizione").Value
											IndirizzoCampo = "" & Rec("Indirizzo").Value
										End If
										Rec.Close()
									End If
								End If

								Dim gf As New GestioneFilesDirectory
								Dim q As Integer = 1
								Dim Body As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\nuova_partita.txt")
								Dim d As DateTime = Convert.ToDateTime(DataOra)
								If Anticipo = 0 Then
									Anticipo = 1
								End If
								Dim qAnticipo As Integer = Anticipo * 60
								Dim OraConvocazione As String = FormatDateTime(d.AddMinutes(-qAnticipo), DateFormat.ShortTime)

								Body = Body.Replace("***DATA***", FormatDateTime(DataOra, DateFormat.LongDate) & " " & FormatDateTime(DataOra, DateFormat.ShortTime))
								Body = Body.Replace("***CAMPO***", sCampo)
								Body = Body.Replace("***ORARIO***", OraConvocazione)
								Body = Body.Replace("***INDIRIZZO***", IndirizzoCampo)
								Body = Body.Replace("***TIPOPARTITA***", tipoPartita)
								Body = Body.Replace("***TELEFONO***", Telefono)
								Body = Body.Replace("***ALLENATORE***", Allenatore)
								Body = Body.Replace("***TELALLENATORE***", TelAllenatore)
								If Lat <> "" And Lon <> "" Then
									Body = Body.Replace("***URLMAPPA***", "https://www.google.it/maps/place/" & Lat & "," & Lon & "z")
								Else
									Body = Body.Replace("***URLMAPPA***", "")
								End If
								Body = Body.Replace("***URLMAPPAAPP***", "https://www.google.it/maps/place/" & LuogoAppuntamento)
								Body = Body.Replace("***REFERENTE***", Referente)
								Body = Body.Replace("***DOAPPUNTAMENTO***", FormatDateTime(DataOraAppuntamento, DateFormat.LongDate) & " " & FormatDateTime(DataOraAppuntamento, DateFormat.ShortTime))
								Body = Body.Replace("***APPUNTAMENTO***", LuogoAppuntamento)
								Dim Mezzo As String = ""
								If MezzoTrasporto = "P" Then
									Mezzo = "Pullman"
								Else
									Mezzo = "Auto propria"
								End If
								Body = Body.Replace("***MEZZO***", Mezzo)

								Dim Paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
								Dim p() As String = Paths.Split(";")
								Dim pathSito As String = p(2).ToUpper().Replace("MULTIMEDIA", "")

								For Each m As String In MailsConvocati
									Dim Oggetto As String = "Nuova partita (" & tipoPartita & ") : "
									Dim Body2 As String = Body
									Dim c() As String = m.Split(";")

									If c(3) = "C" Then
										Body2 = Body2.Replace("***TIPOLOGIA***", "Il giocatore")
									Else
										Body2 = Body2.Replace("***TIPOLOGIA***", "L'allenatore")
									End If

									If Casa = "S" Then
										Oggetto &= Categoria & "-" & Avversario
										Body2 = Body2.Replace("***SQUADRA1***", Categoria)
										Body2 = Body2.Replace("***SQUADRA2***", Avversario)
										Body2 = Body2.Replace("***LUOGO***", "In Casa")
									Else
										Oggetto &= Avversario & "-" & Categoria
										Body2 = Body2.Replace("***SQUADRA1***", Avversario)
										Body2 = Body2.Replace("***SQUADRA2***", Categoria)
										If Casa = "N" Then
											Body2 = Body2.Replace("***LUOGO***", IndirizzoCampo)
										Else
											Body2 = Body2.Replace("***LUOGO***", Campo)
										End If
									End If
									Oggetto &= " " & FormatDateTime(DataOra, DateFormat.LongDate) & " " & FormatDateTime(DataOra, DateFormat.ShortTime)

									Body2 = Body2.Replace("***COGNOME***", c(0))
									Body2 = Body2.Replace("***NOME***", c(1))

									If c(3) <> "A" And c(3) <> "D" Then
										Dim urlSi As String = pathSito & "wsRisposte.asmx/GeneraRisposta?Squadra=" & Squadra & "&Risposta=SI&idPartita=" & idPartita & "&idGiocatore=" & c(4) & "&Tipo=" & c(3)
										Dim urlNo As String = pathSito & "wsRisposte.asmx/GeneraRisposta?Squadra=" & Squadra & "&Risposta=NO&idPartita=" & idPartita & "&idGiocatore=" & c(4) & "&Tipo=" & c(3)

										Body2 = Body2.Replace("***URLPARTECIPO***", urlSi)
										Body2 = Body2.Replace("***URLNONPARTECIPO***", urlNo)
									Else
										Body2 = Body2.Replace("***URLPARTECIPO***", "")
										Body2 = Body2.Replace("***URLNONPARTECIPO***", "")
									End If

									If MandaMail = "S" And (c(3) = "C" Or c(3) = "A" Or c(3) = "D") Then
										Ritorno = ma.SendEmail(Squadra, Mittente, Oggetto, Body2, c(2), {""})
									End If

									If c(3) = "C" Then
										gf.CreaDirectoryDaPercorso(p(0) & "\" & Squadra & "\Convocazioni\Anno" & idAnno & "\Giocatori\")
										gf.CreaAggiornaFile(p(0) & "\" & Squadra & "\Convocazioni\Anno" & idAnno & "\Giocatori\Convocazione_" & idPartita & "_" & c(4) & ".html", Body2)
									End If

									q += 1
								Next
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try

						End If
					End If
				End If

				If Ok Then
					Dim Progressivo As Integer = 1

					For Each C As String In InFormazione.Split(";")
						If C <> "" Then
							' Dim Campi() As String = C.Split(";")
							' Dim idGioc As String = Campi(0)

							If Ok Then
								Sql = "Insert Into InFormazione Values (" &
												" " & idPartita & ", " &
												" " & Progressivo & ", " &
												" " & C & " " &
												")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
									Exit For
								End If

								Progressivo += 1
							Else
								Exit For
							End If
						End If
					Next
				End If

				If Ok Then
					If Dirigenti <> "" Then
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
							For Each C As String In Dirigenti.Split(";")
								If C <> "" Then
									' Dim Campi() As String = C.Split("!")
									Dim idDirigente As String = C.Replace("§", "")

									If Ok Then
										If idDirigente <> "" Then
											Sql = "Insert Into DirigentiPartite Values (" &
												" " & idAnno & ", " &
												" " & idPartita & ", " &
												" " & Progressivo & ", " &
												" " & idDirigente & " " &
												")"
											Ritorno = EsegueSql(Conn, Sql, Connessione)
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
												Exit For
											End If

											Progressivo += 1
										End If
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
				End If

				'If Ok Then
				'	Try
				'		If idArbitro = 0 Or idArbitro = "" Then idArbitro = 1

				'		Sql = "Insert Into ArbitriPartite Values (" &
				'					" " & idAnno & ", " &
				'					" " & idPartita & ", " &
				'					"1, " &
				'					" " & idArbitro & " " &
				'					")"
				'		Ritorno = EsegueSql(Conn, Sql, Connessione)
				'		If Ritorno.Contains(StringaErrore) Then
				'			Ok = False
				'		End If
				'	Catch ex As Exception
				'		Ritorno = StringaErrore & " " & ex.Message
				'		Ok = False
				'	End Try
				'End If

				If Ok Then
					If RigoriPropri <> "" And RigoriAvv.Contains("§") Then
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
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
								End If
							Next
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					If RigoriAvv <> "" Then
						Try
							If RigoriAvv <> "" And RigoriAvv.Contains("§") Then
								Dim a() As String = RigoriAvv.Split("§")

								Sql = "Insert Into RigoriAvversari Values (" &
									" " & idAnno & ", " &
									" " & idPartita & ", " &
									" " & a(0) & ", " &
									" " & a(1) & " " &
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

				If Ok Then
					If EventiPrimoTempo <> "" Then
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
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
										Exit For
									End If
								End If
							Next
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					If EventiSecondoTempo <> "" Then
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
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
										Exit For
									End If
								End If
							Next
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					If EventiTerzoTempo <> "" Then
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
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
										Exit For
									End If
								End If
							Next
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If
				End If

				If Ok Then
					Sql = "Insert Into RisultatiAvversariMinuti Values (" &
									" " & idPartita & ", " &
									"1, " &
									"'" & TempiGoalAvversari1T & "' " &
									")"
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					Else
						Sql = "Insert Into RisultatiAvversariMinuti Values (" &
									" " & idPartita & ", " &
									"2, " &
									"'" & TempiGoalAvversari2T & "' " &
									")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						Else
							Sql = "Insert Into RisultatiAvversariMinuti Values (" &
									" " & idPartita & ", " &
									"3, " &
									"'" & TempiGoalAvversari3T & "' " &
									")"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If
						End If
					End If
				End If

				If Ok Then
					Ritorno = "*"
					Sql = "Commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)

					'If Not Ritorno2.Contains(StringaErrore) Then
					'	If CreaSchedaPartita = "S" Then
					'		Ritorno = CreaHtmlPartita(Squadra, Conn, Connessione, idAnno, idPartita)
					'	End If
					'End If
				Else
					Sql = "Rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaHtmlPerPartita(Squadra As String, idAnno As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Ritorno = CreaHtmlPartita(Squadra, Conn, Connessione, idAnno, idPartita)
			End If
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
						"LEFT JOIN [Generale].[dbo].[TipologiePartite] ON Partite.idTipologia = TipologiePartite.idTipologia) " &
						"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita) " &
						"LEFT JOIN AvversariCoord ON Partite.idAvversario = AvversariCoord.idAvversario) " &
						"LEFT JOIN ArbitriPartite ON (Partite.idPartita = ArbitriPartite.idPartita And Partite.idAnno=ArbitriPartite.idAnno)) " &
						"LEFT JOIN Arbitri ON (Arbitri.idArbitro = ArbitriPartite.idArbitro And ArbitriPartite.idAnno = Arbitri.idAnno)) " &
						"WHERE Partite.idAnno=" & idAnno & "  " &
						"And Partite.idCategoria=" & idCategoria & " Order By DataOra Desc"
					' And Arbitri.idAnno=" & idAnno & " And Partite.Giocata='S'
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

								If MultiMediaPartite <> "" And MultiMediaPartite.Contains("§") Then
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
									"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo = Ruoli.idRuolo) " &
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
						"TipologiePartite.Descrizione AS Tipologia, Allenatori.Cognome+' '+Allenatori.Nome AS Allenatore, Categorie.AnnoCategoria + '-' + Categorie.Descrizione As Categoria, " &
						"CampiAvversari.Indirizzo as CampoIndirizzo, Partite.Casa, Allenatori.idAllenatore, CampiEsterni.Descrizione As CampoEsterno, " &
						"RisultatiAggiuntivi.Tempo1Tempo, RisultatiAggiuntivi.Tempo2Tempo, RisultatiAggiuntivi.Tempo3Tempo, " &
						"CoordinatePartite.Lat, CoordinatePartite.Lon, TempiGoalAvversari.TempiPrimoTempo, TempiGoalAvversari.TempiSecondoTempo, TempiGoalAvversari.TempiTerzoTempo, " &
						"MeteoPartite.Tempo, MeteoPartite.Gradi, MeteoPartite.Umidita, MeteoPartite.Pressione, MeteoPartite.Icona, ArbitriPartite.idArbitro, Arbitri.Cognome + ' ' + Arbitri.Nome As Arbitro, " &
						"Partite.RisultatoATempi, Partite.DataOraAppuntamento, Partite.LuogoAppuntamento, Partite.MezzoTrasporto, Categorie.AnticipoConvocazione, Anni.Indirizzo, Anni.Lat, Anni.Lon, " &
						"Anni.CampoSquadra, Anni.NomePolisportiva, Partite.ShootOut, Partite.Tempi, Partite.PartitaConRigori, PartiteCapitani.idCapitano, " &
						"RisultatiAvversariMinuti1.Minuti As TempiGAvv1, RisultatiAvversariMinuti2.Minuti As TempiGAvv2, RisultatiAvversariMinuti3.Minuti As TempiGAvv3 " &
						"FROM Partite LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita " &
						"LEFT JOIN RisultatiAggiuntivi ON Partite.idPartita = RisultatiAggiuntivi.idPartita " &
						"LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario " &
						"LEFT JOIN [Generale].[dbo].[TipologiePartite] ON Partite.idTipologia = TipologiePartite.idTipologia " &
						"LEFT JOIN Allenatori ON Partite.idAnno = Allenatori.idAnno And Partite.idAllenatore = Allenatori.idAllenatore " &
						"LEFT JOIN CampiAvversari ON SquadreAvversarie.idCampo = CampiAvversari.idCampo " &
						"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita " &
						"LEFT JOIN Categorie ON Partite.idCategoria = Categorie.idCategoria And Categorie.idAnno = Partite.idAnno " &
						"LEFT JOIN CoordinatePartite On Partite.idPartita = CoordinatePartite.idPartita " &
						"LEFT JOIN MeteoPartite On Partite.idPartita = MeteoPartite.idPartita " &
						"LEFT JOIN TempiGoalAvversari On Partite.idPartita = TempiGoalAvversari.idPartita " &
						"LEFT JOIN ArbitriPartite On Partite.idPartita = ArbitriPartite.idPartita And ArbitriPartite.idAnno = Partite.idAnno " &
						"LEFT JOIN Arbitri On ArbitriPartite.idArbitro=Arbitri.idArbitro And ArbitriPartite.idAnno=Arbitri.idAnno " &
						"LEFT JOIN Anni On Partite.idAnno = Anni.idAnno " &
						"LEFT JOIN PartiteCapitani On Partite.idPartita = PartiteCapitani.idPartita " &
						"LEFT JOIN RisultatiAvversariMinuti As RisultatiAvversariMinuti1 On Partite.idPartita = RisultatiAvversariMinuti1.idPartita And RisultatiAvversariMinuti1.idTempo = 1 " &
						"LEFT JOIN RisultatiAvversariMinuti As RisultatiAvversariMinuti2 On Partite.idPartita = RisultatiAvversariMinuti2.idPartita And RisultatiAvversariMinuti2.idTempo = 2 " &
						"LEFT JOIN RisultatiAvversariMinuti As RisultatiAvversariMinuti3 On Partite.idPartita = RisultatiAvversariMinuti3.idPartita And RisultatiAvversariMinuti3.idTempo = 3 " &
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

								'21
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
								Ritorno &= Rec("Icona").Value & ";"

								'32
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
									"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo = Ruoli.idRuolo) " &
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

								'37
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

								'39
								Ritorno &= EventiPrimoTempo & ";"
								Ritorno &= EventiSecondoTempo & ";"
								Ritorno &= EventiTerzoTempo & ";"

								Ritorno &= Rec("Risultato").Value & ";"

								Ritorno &= Rec("DataOraAppuntamento").Value & ";"
								Ritorno &= Rec("LuogoAppuntamento").Value & ";"
								Ritorno &= Rec("MezzoTrasporto").Value & ";"
								Ritorno &= Rec("AnticipoConvocazione").Value & ";"
								Ritorno &= Rec("Indirizzo").Value & ";"
								Ritorno &= Rec("Lat").Value & ";"
								Ritorno &= Rec("Lon").Value & ";"
								Ritorno &= Rec("CampoSquadra").Value & ";"
								Ritorno &= Rec("NomePolisportiva").Value & ";"
								Ritorno &= Rec("ShootOut").Value & ";"
								Ritorno &= Rec("Tempi").Value & ";"
								Ritorno &= Rec("PartitaConRigori").Value & ";"
								Ritorno &= Rec("idCapitano").Value & ";"
								Dim ga1 As String = IIf(Rec("TempiGAvv1").Value Is DBNull.Value, "", Rec("TempiGAvv1").Value)
								Dim ga2 As String = IIf(Rec("TempiGAvv2").Value Is DBNull.Value, "", Rec("TempiGAvv2").Value)
								Dim ga3 As String = IIf(Rec("TempiGAvv3").Value Is DBNull.Value, "", Rec("TempiGAvv3").Value)
								Ritorno &= ga1.Replace(";", "%") & ";"
								Ritorno &= ga2.Replace(";", "%") & ";"
								Ritorno &= ga3.Replace(";", "%") & ";"
								Ritorno &= "§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()
						Ritorno &= "|"

						Sql = "Select * From (Select idTempo, Progressivo, RisultatiAggiuntiviMarcatori.idGiocatore, Minuto, Cognome, Nome, Ruoli.Descrizione As Ruolo, NumeroMaglia, Rigore " &
							"FROM RisultatiAggiuntiviMarcatori " &
							"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
							"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo = Ruoli.idRuolo " &
							"Where RisultatiAggiuntiviMarcatori.idPartita=" & idPartita & " And Giocatori.idAnno=" & idAnno & " " &
							"Union All " &
							"Select idTempo, Progressivo, -1, Minuto, 'Autorete' As Cognome, '' As Nome, '' As Ruolo, 999 As NumeroMaglia, 'N' As Rigore FROM RisultatiAggiuntiviMarcatori " &
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
										Rec("Rigore").Value.ToString & ";" &
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
							"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo = Ruoli.idRuolo) " &
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

						' Convocati
						Sql = "SELECT idProgressivo, Convocati.idGiocatore, Cognome, Nome, Ruoli.idRuolo, Ruoli.Descrizione As Ruolo, NumeroMaglia, ConvocatiPartiteRisposte.Risposta " &
							"FROM Convocati " &
							"Left Join Giocatori On Convocati.idGiocatore = Giocatori.idGiocatore " &
							"Left Join ConvocatiPartiteRisposte On ConvocatiPartiteRisposte.idGiocatore = Convocati.idGiocatore And ConvocatiPartiteRisposte.idPartita = Convocati.idPartita " &
							"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo = Ruoli.idRuolo " &
							"Where Convocati.idPartita=" & idPartita & " And Giocatori.idAnno=" & idAnno & " Order By Cognome, Nome"
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
										Rec("Risposta").Value.ToString & ";" &
										"§"

									Rec.MoveNext()
								Loop
								Ritorno &= "|"
							End If
							Rec.Close()
						End If

						' In Formazione
						Sql = "SELECT idProgressivo, InFormazione.idGiocatore, Cognome, Nome, Ruoli.idRuolo, Ruoli.Descrizione As Ruolo, NumeroMaglia " &
							"FROM InFormazione " &
							"Left Join Giocatori On InFormazione.idGiocatore = Giocatori.idGiocatore " &
							"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo = Ruoli.idRuolo " &
							"Where InFormazione.idPartita=" & idPartita & " And Giocatori.idAnno=" & idAnno & " Order By idProgressivo"
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

							' Esiste PDF
							Dim gf As New GestioneFilesDirectory
							Dim sIdPartita As String = idPartita.Trim
							For i As Integer = sIdPartita.Length - 1 To 3
								sIdPartita = "0" & sIdPartita
							Next
							Dim paths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
							Dim P() As String = paths.Split(";")
							If Strings.Right(P(0), 1) <> "\" Then
								P(0) &= "\"
							End If
							Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
							If Strings.Right(P(2), 1) <> "/" Then
								P(2) &= "/"
							End If
							Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
							Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Partite\Anno" & idAnno & "\" & sIdPartita & "\" & idPartita & ".pdf"
							'Return NomeFileFinalePDF

							If File.Exists(NomeFileFinalePDF) Then
								Ritorno &= NomeFileFinalePDF.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/") & "|"
							Else
								Ritorno &= "|"
							End If
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
				Dim idPartita1 As Integer
				Dim idPartita2 As Integer

				Try
					Sql = "SELECT Max(idPartita)+1 FROM Partite"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec(0).Value Is DBNull.Value Then
							idPartita1 = 1
						Else
							idPartita1 = Rec(0).Value
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try


				Try
					Sql = "SELECT Max(idPartita)+1 FROM CalendarioPartite"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun progressivo partita rilevato"
						Else
							If Rec(0).Value Is DBNull.Value Then
								idPartita2 = 1
							Else
								idPartita2 = Rec(0).Value
							End If
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				If idPartita1 >= idPartita2 Then
					Ritorno = idPartita1
				Else
					Ritorno = idPartita2
				End If

			End If

			Conn.Close()
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
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Sql = "SELECT Partite.Casa, Partite.idPartita, Partite.DataOra, Categorie.Descrizione As Categoria, SquadreAvversarie.Descrizione As Avversario, '' + CampiAvversari.Descrizione As Campo, " &
					"CampiAvversari.Indirizzo, '' + CampiEsterni.Descrizione As CampoEsterno, Allenatori.Cognome + ' ' + Allenatori.Nome As Mister, Allenatori.Telefono, Partite.OraConv, " &
					"Anni.CampoSquadra, Anni.Indirizzo As IndirizzoCasa, Categorie.AnticipoConvocazione, Partite.DataOraAppuntamento, Partite.LuogoAppuntamento, Partite.MezzoTrasporto, " &
					"Anni.NomePolisportiva, Anni.NomeSquadra, Anni.Indirizzo " &
					"FROM (((((Partite LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) " &
					"LEFT JOIN CampiAvversari ON SquadreAvversarie.idCampo = CampiAvversari.idCampo) " &
					"LEFT JOIN Categorie ON (Partite.idAnno = Categorie.idAnno) And (Partite.idCategoria = Categorie.idCategoria)) " &
					"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita) " &
					"LEFT JOIN Allenatori ON (Partite.idAllenatore = Allenatori.idAllenatore) And (Partite.idCategoria = Allenatori.idCategoria) And (Partite.idAnno = Allenatori.idAnno)) " &
					"LEFT JOIN Anni On Partite.idAnno = Anni.idAnno " &
					"WHERE Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						If Not Rec("DataOra").Value Is DBNull.Value Then
							Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
							Dim pp As String = paths
							pp = pp.Replace(vbCrLf, "")
							If Strings.Right(pp, 1) <> "\" Then
								pp = pp & "\"
							End If

							paths = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
							Dim p() As String = paths.Split(";")
							If Strings.Right(p(0), 1) <> "\" Then
								p(0) = p(0) & "\"
							End If
							p(0) = p(0).Replace(vbCrLf, "")
							If Strings.Right(p(2), 1) <> "/" Then
								p(2) = p(2) & "/"
							End If
							p(2) = p(2).Replace(vbCrLf, "")

							' Dim Anticipo As Single = ("" & Rec("AnticipoConvocazione").Value).replace(",", ".")
							'If Anticipo = 0 Then
							'	Anticipo = 1
							'End If
							Dim Filetto As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\base_convocazioni.txt")
							Dim Datella As Date = Rec("DataOra").Value
							'Dim DatellaConv As Date = Datella.AddHours(-Anticipo)

							Filetto = Filetto.Replace("***PARTITA***", idPartita)

							Filetto = Filetto.Replace("***SQUADRA***", Rec("Categoria").Value)

							'Dim url As String = p(2) & Rec("NomeSquadra").Value.ToString.Replace(" ", "_") & "/Societa/" & idAnno & "_1.kgb"
							'Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
							'Dim urlConv As String = ""
							' C:\IIS\GestioneCampionato\CalcioImages\Morti_De_Sonno_Fc\Societa
							Dim fileUrlOrig As String = p(2) & Rec("NomeSquadra").Value.ToString.Replace(" ", "_") & "/Societa/Societa_1.png"
							' If File.Exists(fileUrlOrig) Then
							'	Dim fileUrlConv As String = pp & "Appoggio/" & Esten & ".jpg"
							'	Dim c As New CriptaFiles
							'	c.DecryptFile(CryptPasswordString, fileUrlOrig, fileUrlConv)
							'urlConv = fileUrlOrig ' p(2) & "Appoggio/" & Esten & ".jpg"
							'End If

							Filetto = Filetto.Replace("***URL LOGO***", fileUrlOrig)

							Filetto = Filetto.Replace("***NOME POLISPORTIVA***", Rec("NomePolisportiva").Value)
							Filetto = Filetto.Replace("***INDIRIZZO POLISPORTIVA***", Rec("Indirizzo").Value)

							Filetto = Filetto.Replace("***GARA***", Rec("Categoria").Value & " - " & Rec("Avversario").Value)
							Filetto = Filetto.Replace("***DATA***", Format(Datella.Day, "00") & "/" & Format(Datella.Month, "00") & "/" & Datella.Year)
							If Not Rec("CampoEsterno").Value Is DBNull.Value And "" & Rec("CampoEsterno").Value <> "" Then
								Filetto = Filetto.Replace("***CAMPO***", Rec("CampoEsterno").Value)
								Filetto = Filetto.Replace("***INDIRIZZO***", "")
							Else
								If Rec("Casa").Value = "S" Then
									Filetto = Filetto.Replace("***CAMPO***", Rec("CampoSquadra").Value)
									Filetto = Filetto.Replace("***INDIRIZZO***", Rec("IndirizzoCasa").Value)
								Else
									Filetto = Filetto.Replace("***CAMPO***", Rec("Campo").Value)
									Filetto = Filetto.Replace("***INDIRIZZO***", Rec("Indirizzo").Value)
								End If
							End If

							Dim Appuntamento As String = "" & Rec("DataOraAppuntamento").Value
							If Appuntamento <> "" Then
								Appuntamento = Mid(Appuntamento, Appuntamento.IndexOf(" ") + 1, Appuntamento.Length)
							End If
							Dim OraConv As String = "" & Rec("DataOra").Value
							If OraConv <> "" Then
								OraConv = Mid(OraConv, OraConv.IndexOf(" ") + 1, OraConv.Length)
							End If
							Filetto = Filetto.Replace("***ORARIO1***", Appuntamento)
							Filetto = Filetto.Replace("***ORARIO2***", OraConv)

							Filetto = Filetto.Replace("***MISTER***", "" & Rec("Mister").Value)
							Filetto = Filetto.Replace("***CELL***", "" & Rec("Telefono").Value)

							'Filetto = Filetto.Replace("***DOAPPUNTAMENTO***", Rec("DataOraAppuntamento").Value)
							'Filetto = Filetto.Replace("***APPUNTAMENTO***", Rec("LuogoAppuntamento").Value)

							'Dim Mezzo As String = ""
							'If "" & Rec("MezzoTrasporto").Value = "P" Then
							'	Mezzo = "Pullman"
							'Else
							'	Mezzo = "Auto propria"
							'End If
							'Filetto = Filetto.Replace("***MEZZO***", mezzo)

							Rec.Close()

							Dim Convocati As String = ""

							Sql = "SELECT Giocatori.Cognome+' '+Giocatori.Nome AS Giocatore, Ruoli.idRuolo " &
								"FROM (Convocati LEFT JOIN Giocatori ON Convocati.idGiocatore = Giocatori.idGiocatore) LEFT JOIN [Generale].[dbo].[Ruoli] ON Giocatori.idRuolo = Ruoli.idRuolo " &
								"WHERE Convocati.idPartita=" & idPartita & " AND Giocatori.idAnno=" & idAnno & " " &
								"ORDER BY Ruoli.idRuolo, Giocatori.Cognome, Giocatori.Nome"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								Dim Giocatori As List(Of String) = New List(Of String)

								Do Until Rec.Eof
									Giocatori.Add("" & Rec("Giocatore").Value)

									Rec.MoveNext
								Loop
								Rec.Close

								Convocati &= "<table style=""width: 100%;"" cellpadding=""0px"" cellspacing=""0px"">"
								For i As Integer = 0 To 11
									Dim Riga As String = "<tr>"

									Riga &= "<td style=""width: 10%;"" class=""adestra"">"
									Riga &= "<span class=""titolo3"">" & i + 1 & "</span>"
									Riga &= "</td>"

									Riga &= "<td style=""width:10px;"">"
									Riga &= "&nbsp;"
									Riga &= "</td>"

									Riga &= "<td  style=""width: 35%;"">"
									If i < Giocatori.Count Then
										Riga &= "<span class=""titolo3"">" & Giocatori.Item(i) & "</span>"
									Else
										Riga &= "<span class=""titolo3"">&nbsp;</span>"
									End If
									Riga &= "</td>"

									Riga &= "<td style=""width: 10%;"" class=""adestra"">"
									Riga &= "<span class=""titolo3"">" & i + 11 & "</span>"
									Riga &= "</td>"

									Riga &= "<td style=""width:10px;"">"
									Riga &= "&nbsp;"
									Riga &= "</td>"

									Riga &= "<td style=""width: 35%;"">"
									If i + 12 < Giocatori.Count Then
										Riga &= "<span class=""titolo3"">" & Giocatori.Item(i + 12) & "</span>"
									Else
										Riga &= "<span class=""titolo3"">&nbsp;</span>"
									End If
									Riga &= "</td>"
									Riga &= "</tr>"

									Convocati &= Riga
								Next
								Convocati &= "</table>"
							End If
							Filetto = Filetto.Replace("***CONVOCATI***", Convocati)

							Dim pathFileAgg As String = p(0) & Squadra & "\Scheletri\testo_convocazioni.txt"
							If Not File.Exists(pathFileAgg) Then
								pathFileAgg = Server.MapPath(".") & "\Scheletri\testo_convocazioni.txt"
							End If
							Dim testo As String = gf.LeggeFileIntero(pathFileAgg)
							Filetto = Filetto.Replace("***TESTO AGGIUNTIVO***", testo)

							Dim Dirigenti As String = "" ' "<Table style=""width: 100%;"" cellpadding=""0px"" cellspacing=""0px"">"
							Sql = "Select Cognome, Nome, Telefono From DirigentiPartite A " &
								"Left Join Dirigenti B On A.idDirigente = B.idDirigente " &
								"Where A.idPartita = " & idPartita & " And Eliminato = 'N' " &
								"Order By Progressivo"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								' Dirigenti &= "<tr style=""border: 1px solid #999""><th><span style=""font-family: Arial; font-size: 16px;"">Dirigente</span></th><th><span style=""font-family: Arial; font-size: 16px;"">Telefono</span></th></tr>"
								Do Until Rec.Eof
									Dirigenti &= Rec("Cognome").Value & " " & Rec("Nome").Value & "<br />" & Rec("Telefono").Value & "<br />"

									Rec.MoveNext
								Loop
								Rec.Close
							End If
							'Dirigenti &= "</table>"
							Filetto = Filetto.Replace("***DIRIGENTI***", Dirigenti)

							Dim path1 As String = p(0) & "\" & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".html"
							Dim pathPdf As String = p(0) & "\" & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".pdf"
							Dim pathLog As String = p(0) & "\" & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".log"

							gf.CreaDirectoryDaPercorso(p(0) & "\" & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\")
							gf.EliminaFileFisico(path1)
							gf.EliminaFileFisico(pathPdf)
							gf.CreaAggiornaFile(path1, Filetto)

							Ritorno = "Allegati/" + Squadra & "/Convocazioni/Anno" & idAnno & "/Partite/Partita_" & idPartita & ".pdf"

							Dim pp2 As New pdfGest
							Dim Ritorno2 As String = pp2.ConverteHTMLInPDF(path1, pathPdf, pathLog)
							If Ritorno = "*" Then
								Ritorno = Ritorno2
							End If
						Else
							Ritorno = StringaErrore & " Data non valida"
						End If
					Else
						Ritorno = StringaErrore & " Nessun dato rilevato"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaPartitaGEN(Squadra As String, idAnno As String, idPartita As String) As String
		Return EliminaPartita(Squadra, idAnno, idPartita)
	End Function

	<WebMethod()>
	Public Function CreaFoglioConvocazionePDF(Squadra As String, idAnno As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathLog As String = p(0) & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".log"
		Dim path1 As String = p(0) & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".html"
		Dim pathPdf As String = p(0) & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".pdf"
		gf.CreaDirectoryDaPercorso(pathLog)
		gf.CreaDirectoryDaPercorso(path1)
		gf.CreaDirectoryDaPercorso(pathPdf)
		Dim pp As New pdfGest
		Ritorno = pp.ConverteHTMLInPDF(path1, pathPdf, pathLog)
		If Ritorno = "*" Then
			Ritorno = pathPdf
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InviaFoglioConvocazionePDF(Squadra As String, idAnno As String, idPartita As String, Mittente As String) As String
		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim path1 As String = p(0) & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".html"
		Dim pathLog As String = p(0) & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".log"
		Dim pathPdf As String = p(0) & Squadra & "\Convocazioni\Anno" & idAnno & "\Partite\Partita_" & idPartita & ".pdf"
		Dim IndirizzoWS As String = p(2)
		p(2) = p(2).Replace(vbCrLf, "")
		p(2) = p(2).Replace("Multimedia", "")
		If Strings.Right(IndirizzoWS, 1) <> "/" Then
			IndirizzoWS &= "/"
		End If

		If Not File.Exists(pathPdf) Then
			Dim pp As New pdfGest
			Ritorno = pp.ConverteHTMLInPDF(path1, pathPdf, pathLog)
		Else
			Ritorno = "*"
		End If

		If Ritorno = "*" Then
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
					Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
					Dim Sql As String = ""

					Dim Tipologia As String = ""
					Dim Categoria As String = ""
					Dim Avversario As String = ""
					Dim DataOra As String = ""
					Dim DataOraAppuntamento As String = ""
					Dim LuogoAppuntamento As String = ""
					Dim Allenatore As String = ""
					Dim Dirigenti As String = ""
					Dim Casa As String = ""
					Dim NomePolisportiva As String = ""

					Sql = "Select isnull(B.Descrizione, '') As Tipologia, C.Descrizione As Categoria, D.Descrizione As Avversario, DataOra, DataOraAppuntamento As Appuntamento, LuogoAppuntamento, " &
						"E.Cognome + ' ' + E.Nome As Allenatore, Casa " &
						"From Partite A " &
						"Left Join [Generale].[dbo].[TipologiePartite] B On A.idTipologia = B.idTipologia " &
						"Left Join Categorie C On A.idCategoria = C.idCategoria  " &
						"Left Join SquadreAvversarie D On A.idAvversario = D.idAvversario  " &
						"Left Join Allenatori E On A.idAllenatore = E.idAllenatore " &
						"Where idPartita = " & idPartita
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If Not Rec.Eof Then
						Casa = Rec("Casa").Value
						Tipologia = Rec("Tipologia").Value
						Categoria = Rec("Categoria").Value
						Avversario = Rec("Avversario").Value
						DataOra = Rec("DataOra").Value
						DataOraAppuntamento = Rec("Appuntamento").Value
						LuogoAppuntamento = Rec("LuogoAppuntamento").Value
						Allenatore = Rec("Allenatore").Value
					Else
						Return "ERROR: Nessun dato trovato per la partita"
					End If

					Dim Oggetto As String = "Nuova partita (" & Tipologia & ") : "

					If Casa = "S" Then
						Oggetto &= Categoria & "-" & Avversario & " In Casa"
					Else
						Oggetto &= Avversario & "-" & Categoria
						Sql = "Select B.Descrizione As Campo, C.Descrizione As CampoEsterno From Partite A " &
							"Left Join CampiAvversari B On A.idCampo = B.idCampo " &
							"Left Join CampiEsterni C On A.idPartita = C.idPartita " &
							"Where A.idPartita = " & idPartita
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Rec.Eof Then
							Oggetto = " Campo esterno"
						Else
							If Casa = "N" Then
								Oggetto &= " " & Rec("Campo").Value
							Else
								Oggetto &= " " & Rec("CampoEsterno").Value
							End If
						End If
					End If
					Oggetto &= " " & FormatDateTime(DataOra, DateFormat.LongDate) & " " & FormatDateTime(DataOra, DateFormat.ShortTime)

					'Sql = "Select * From DirigentiPartite A " &
					'	"Left Join Dirigenti B On A.idDirigente = B.idDirigente " &
					'	"Where idPartita = " & idPartita
					'Rec = LeggeQuery(Conn, Sql, Connessione)
					'Dirigenti = ""
					'Do Until Rec.Eof
					'	Dirigenti = Rec("Cognome").Value & " " & Rec("Nome").Value & ", "

					'	Rec.MoveNext
					'Loop
					'If Dirigenti.Length > 0 Then
					'	Dirigenti = Mid(Dirigenti, 1, Len(Dirigenti) - 2)
					'End If

					Sql = "Select * From Anni"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If Not Rec.Eof Then
						NomePolisportiva = Rec("NomePolisportiva").Value
					End If

					Sql = "Select Distinct EMail, id, Tipo From ( " &
						"Select EMail, B.idDirigente As id, 'Dirigente' As Tipo From Partite A " &
						"Left Join DirigentiPartite B On A.idPartita = B.idPartita  " &
						"Left Join Dirigenti C On B.idDirigente = C.idDirigente " &
						"Where A.idPartita = " & idPartita & " " &
						"Union All " &
						"Select EMail, A.idGiocatore As id, 'Giocatore Maggiorenne' As Tipo From Convocati A " &
						"Left Join Giocatori B On A.idGiocatore = B.idGiocatore  " &
						"Where idPartita = " & idPartita & " And Maggiorenne = 'S'  " &
						"Union All  " &
						"Select C.Mail As Email, A.idGiocatore As id, 'Giocatore' As Tipo From Convocati A " &
						"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
						"Left Join GiocatoriMails C On A.idGiocatore = C.idGiocatore " &
						"Where idPartita = " & idPartita & " And Maggiorenne = 'N' And Attiva = 'S' " &
						"Union All " &
						"Select EMail, A.idAllenatore As id, 'Allenatore' As Tipo From Partite A " &
						"Left Join Allenatori B On A.idAllenatore = B.idAllenatore  " &
						"Where idPartita = " & idPartita & " " &
						") A Where EMail <> '' And EMail Is Not Null"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						'Dim Giocatori As List(Of String) = New List(Of String)
						'Dim pathTemplate As String = p(0) & Squadra & "\Scheletri\mail_convocazione.txt"
						'If Not File.Exists(pathTemplate) Then
						'	pathTemplate = Server.MapPath(".") & "\Scheletri\mail_convocazione.txt"
						'End If
						'Dim templ As String = gf.LeggeFileIntero(pathTemplate)
						Dim ma As New mail
						'Dim Oggetto As String = "Convocazione nuova partita: "

						Do Until Rec.Eof
							If Not Rec(0).Value Is DBNull.Value Then
								Dim Body2 As String = ""

								'Body2 = Body2.Replace("***Tipologia Partita***", Tipologia)
								'Body2 = Body2.Replace("***categoria***", Categoria)
								'Body2 = Body2.Replace("***Avversario***", Avversario)
								'Body2 = Body2.Replace("***Data Partita***", DataOra)
								'Body2 = Body2.Replace("***Ora Partita***", "")
								'Body2 = Body2.Replace("***Data Appuntamento***", DataOraAppuntamento)
								'Body2 = Body2.Replace("***Ora Appuntamento***", "")
								'Body2 = Body2.Replace("***Luogo Appuntamento***", LuogoAppuntamento)
								'Body2 = Body2.Replace("***Allenatore***", Allenatore)
								'Body2 = Body2.Replace("***Dirigenti***", Dirigenti)
								'Body2 = Body2.Replace("***nome societa menu settaggi***", NomePolisportiva)

								' Giocatori.Add(Rec(0).Value)
								Select Case Rec(2).Value
									Case "Dirigente"
										Body2 = "Lei è convocato per la partita in oggetto.<br /><br />Saluti<br />" & NomePolisportiva
										'Body2 = Body2.Replace("***cognome menu anagrafica3***", "***")
										'Body2 = Body2.Replace("***Nome menu anagrafica3***", "***")

										'Body2 = Body2.Replace("***Partecipo***", "<div style=""width: 100%""><div style=""background-color: green; color: white; width:50%; float: left; text-align: center;"">---</div>")
										'Body2 = Body2.Replace("***Non Posso***", "<div style=""background-color: red; color: white; width: 50%; float: left; text-align: center;"">---</div></div>")

										Ritorno = ma.SendEmail(Squadra, Mittente, Oggetto, Body2, Rec(0).Value, {pathPdf})
									Case "Giocatore Maggiorenne", "Giocatore"
										Sql = "Select * From Giocatori Where idGiocatore = " & Rec(1).Value
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If Not Rec2.eof Then
											Body2 = "Il giocatore " & Rec2("Cognome").Value & " " & Rec2("Nome").Value & " è convocato per la partita in oggetto.<br /><br />Saluti<br />" & NomePolisportiva
											'Body2 = Body2.Replace("***cognome menu anagrafica3***", Rec2("Cognome").Value)
											'Body2 = Body2.Replace("***Nome menu anagrafica3***", Rec2("Nome").Value)

											'Dim IndirizzoOK As String = IndirizzoWS & "wsRisposte.asmx/GeneraRisposta?Squadra=" & Squadra & "&Risposta=SI&idPartita=" & idPartita & "&idGiocatore=" & Rec(1).Value
											'Dim IndirizzoKO As String = IndirizzoWS & "wsRisposte.asmx/GeneraRisposta?Squadra=" & Squadra & "&Risposta=NO&idPartita=" & idPartita & "&idGiocatore=" & Rec(1).Value

											'Body2 = Body2.Replace("***Partecipo***", "<div style=""width: 100%""><div style=""background-color: green; color: white; width:50%; float: left; text-align: center;""><a href=""" & IndirizzoOK & """>Partecipo</a></div>")
											'Body2 = Body2.Replace("***Non Posso***", "<div style=""background-color: red; color: white; width: 50%; float: left; text-align: center;""><a href=""" & IndirizzoKO & """>Non Posso</a></div></div>")

											Ritorno = ma.SendEmail(Squadra, Mittente, Oggetto, Body2, Rec(0).Value, {pathPdf})
										End If
									Case "Allenatore"
										Body2 = "Lei è convocato per la partita in oggetto.<br /><br />Saluti<br />" & NomePolisportiva
										'Body2 = Body2.Replace("***cognome menu anagrafica3***", "***")
										'Body2 = Body2.Replace("***Nome menu anagrafica3***", "***")

										'Body2 = Body2.Replace("***Partecipo***", "<div style=""width: 100%""><div style=""background-color: green; color: white; width:50%; float: left; text-align: center;"">---</div>")
										'Body2 = Body2.Replace("***Non Posso***", "<div style=""background-color: red; color: white; width: 50%; float: left; text-align: center;"">---</div></div>")

										Ritorno = ma.SendEmail(Squadra, Mittente, Oggetto, Body2, Rec(0).Value, {pathPdf})
								End Select
							End If

							Rec.MoveNext
						Loop
						Rec.Close

					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiungeSostituzioni(Squadra As String, idPartita As String, Dati As String) As String
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
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
				If Ritorno <> "*" Then
					Ok = False
				End If

				If Ok Then
					Sql = "Delete From PartiteSostituzioni Where idPartita = " & idPartita
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno <> "*" Then
						Ok = False
					End If
				End If

				If Ok Then
					Dim campi() As String = Dati.Split("§")
					For Each c As String In campi
						If c <> "" Then
							Dim cc() As String = c.Split(";")
							Dim idSostituito As String = cc(0)
							Dim idEntrante As String = cc(1)
							Dim Tempo As String = cc(2)
							Dim Minuto As String = cc(3)
							Dim Progressivo As Integer

							Sql = "Select Max(Progressivo)+1 From PartiteSostituzioni Where idPartita=" & idPartita & " And Tempo=" & Tempo
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
								Rec.Close
							End If

							If Ok Then
								Sql = "Insert Into PartiteSostituzioni Values (" & idPartita & ", " & Tempo & ", " & Progressivo & ", " & idSostituito & ", " & idEntrante & ", " & Minuto & ")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno <> "*" Then
									Ok = False
									Exit For
								End If
							End If
						End If
					Next
				End If

				If Ok Then
					Ritorno = "*"
					Sql = "Commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "Rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaSostituzioni(Squadra As String, idPartita As String) As String
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
				Dim Sql As String = "Select * From PartiteSostituzioni Where idPartita=" & idPartita & " Order By Tempo, Minuto"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("idSostituito").Value & ";" & Rec("idEntrante").Value & ";" & Rec("Tempo").Value & ";" & Rec("Minuto").Value & "§"

						Rec.MoveNext
					Loop
					Rec.Close
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaSostituzione(Squadra As String, idPartita As String, Progressivo As String) As String
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
				Dim Sql As String = "Delete From PartiteSostituzioni Where idPartita=" & idPartita & " And Progressivo=" & Progressivo
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function
End Class