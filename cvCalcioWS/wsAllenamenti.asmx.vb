Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO

<System.Web.Services.WebService(Namespace:="http://cvcalcio_allti.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsAllenamenti
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function SalvaAllenamenti(Squadra As String, idAnno As String, idCategoria As String, Data As String, Ora As String, Giocatori As String, OraFine As String) As String
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
					Sql = "Delete From Allenamenti Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Datella='" & Data & "' And Orella='" & Ora & "'"
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						Dim sGiocatori() As String = Giocatori.Split(";")
						Dim Progressivo As Integer = 0

						For Each s As String In sGiocatori
							If s <> "" Then
								Sql = "Insert Into Allenamenti Values (" &
									" " & idAnno & ", " &
									" " & idCategoria & ", " &
									"'" & Data & "', " &
									"'" & Ora & "', " &
									" " & Progressivo & ", " &
									" " & s & ", " &
									"'" & OraFine & "' " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Sql = "rollback"
									Dim Ritorno3 As String = EsegueSql(Conn, Sql, Connessione)
									Exit For
								End If

								Progressivo += 1
							End If
						Next
						If Not Ritorno.Contains(StringaErrore) Then
							Ritorno = "*"

							Sql = "commit"
							Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
						End If
					Else
						Sql = "rollback"
						Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
					End If
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
	Public Function RitornaOreAllenamenti(Squadra As String, ByVal idAnno As String, idCategoria As String, Data As String) As String
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
					Dim sq() As String = Squadra.Split("_")
					Dim codSquadra As Integer = Val(sq(1))
					Dim NomeSquadra As String = ""
					Dim Ok As Boolean = True

					Sql = "Select Orella From Allenamenti Where Datella = '" & Data & "' And idCategoria = " & idCategoria & " Group By Orella"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun allenamento rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("Orella").Value & ";"

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
	Public Function RitornaAllenamentiCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String, Data As String, Ora As String, Stampa As String) As String
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
					Dim sq() As String = Squadra.Split("_")
					Dim codSquadra As Integer = Val(sq(1))
					Dim NomeSquadra As String = ""
					Dim Ok As Boolean = True

					If Stampa = "S" Then
						Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & codSquadra
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Nessuna squadra rilevata"
								Ok = False
							Else
								NomeSquadra = Rec("Descrizione").Value
							End If
							Rec.Close()
						End If
					End If

					If Ok Then
						Dim NomeCategoria As String = ""

						If Stampa = "S" Then
							Sql = "Select * From Categorie Where idCategoria = " & idCategoria
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
							Else
								If Rec.Eof Then
									Ritorno = StringaErrore & " Nessuna categoria rilevata"
									Ok = False
								Else
									NomeCategoria = Rec("Descrizione").Value
								End If
								Rec.Close()
							End If
						End If

						If Ok Then
							Dim codiciGiocatore As String = ""
							Dim oraInizio As String = ""
							Dim oraFine As String = ""

							Sql = "SELECT Giocatori.idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
								"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, idCategoria2, Categorie.Descrizione As Categoria2, idCategoria3, " &
								"Cat3.Descrizione As Categoria3, Cat1.Descrizione As Categoria1, Allenamenti.Orella, Allenamenti.OrellaFine " &
								"FROM (((((Allenamenti LEFT JOIN Giocatori ON (Allenamenti.idGiocatore = Giocatori.idGiocatore) And (Allenamenti.idAnno = Giocatori.idAnno))) " &
								"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo) " &
								"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria2 And Categorie.idAnno=Giocatori.idAnno) " &
								"Left Join Categorie As Cat3 On Cat3.idCategoria=Giocatori.idCategoria3 And Cat3.idAnno=Giocatori.idAnno) " &
								"Left Join Categorie As Cat1 On Cat1.idCategoria=Giocatori.idCategoria And Cat1.idAnno=Giocatori.idAnno " &
								"Where Giocatori.Eliminato='N' And Allenamenti.idAnno=" & idAnno & " And Allenamenti.idCategoria=" & idCategoria & " And Datella='" & Data & "' And Orella='" & Ora & "' " &
								"Order By Ruoli.idRuolo, Cognome, Nome"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = StringaErrore & " Nessun allenamento rilevato"
								Else
									If Stampa = "N" Then
										Ritorno = ""
									Else
										Ritorno = "<table style=""width: 100%;"">"
										Ritorno &= "<tr><th>Giocatori Presenti</th></tr>"
									End If

									Dim q As Integer = 0

									Do Until Rec.Eof
										If Stampa = "N" Then
											Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
												Rec("idR").Value.ToString & ";" &
												Rec("Cognome").Value.ToString.Trim & ";" &
												Rec("Nome").Value.ToString.Trim & ";" &
												Rec("Descrizione").Value.ToString.Trim & ";" &
												Rec("EMail").Value.ToString.Trim & ";" &
												Rec("Telefono").Value.ToString.Trim & ";" &
												Rec("Soprannome").Value.ToString.Trim & ";" &
												Rec("DataDiNascita").Value.ToString & ";" &
												Rec("Indirizzo").Value.ToString.Trim & ";" &
												Rec("CodFiscale").Value.ToString.Trim & ";" &
												Rec("Maschio").Value.ToString.Trim & ";" &
												Rec("Citta").Value.ToString.Trim & ";" &
												Rec("Matricola").Value.ToString.Trim & ";" &
												Rec("NumeroMaglia").Value.ToString.Trim & ";" &
												Rec("idCategoria").Value.ToString & ";" &
												Rec("idCategoria2").Value.ToString & ";" &
												Rec("Categoria2").Value.ToString & ";" &
												Rec("idCategoria3").Value.ToString & ";" &
												Rec("Categoria3").Value.ToString & ";" &
												Rec("Categoria1").Value.ToString & ";" &
												"§"
										Else
											codiciGiocatore &= Rec("idGiocatore").Value & ","
											oraInizio = Rec("Orella").Value
											oraFine = Rec("OrellaFine").Value

											q += 1
											Dim conta As String = Format(q, "00")

											Ritorno &= "<tr>"
											' Ritorno &= "<td>" & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
											Ritorno &= "<td style=""padding-left: 50px;"">" & conta & " - " & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
											Ritorno &= "</tr>"
										End If

										Rec.MoveNext()
									Loop
								End If
								Rec.Close()

								If Stampa <> "N" Then
									Ritorno &= "</table>"
									Ritorno &= "<hr />"

									If codiciGiocatore.Length > 0 Then
										codiciGiocatore = Mid(codiciGiocatore, 1, codiciGiocatore.Length - 1)
									End If

									Sql = "Select * From Giocatori " &
										"Where CharIndex('" & idCategoria & "-', Categorie) > 0 " &
										"And idGiocatore Not In (" & codiciGiocatore & ")"
									Rec = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										Dim q As Integer = 0

										Ritorno &= "<table style=""width: 100%;"">"
										Ritorno &= "<tr><th>Giocatori Assenti</th></tr>"
										Do Until Rec.Eof
											q += 1
											Dim conta As String = Format(q, "00")

											Ritorno &= "<tr>"
											Ritorno &= "<td style=""padding-left: 50px;"">" & conta & " - " & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
											Ritorno &= "</tr>"

											Rec.MoveNext()
										Loop
										Rec.Close()
										Ritorno &= "</table>"
									End If

									Dim gf As New GestioneFilesDirectory
									Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_allenamenti.txt")

									filetto = filetto.Replace("***TITOLO***", "Allenamenti " & NomeCategoria & "<br />" & Data & " " & oraInizio & "-" & oraFine)
									filetto = filetto.Replace("***DATI***", Ritorno)
									filetto = filetto.Replace("***NOME SQUADRA***", "<br /><br />" & NomeSquadra)

									Dim multimediaPaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
									Dim mmPaths() As String = multimediaPaths.Split(";")
									mmPaths(2) = mmPaths(2).Replace(vbCrLf, "")
									If Strings.Right(mmPaths(2), 1) <> "/" Then
										mmPaths(2) &= "/"
									End If

									Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
									filePaths = filePaths.Replace(vbCrLf, "")
									If Strings.Right(filePaths, 1) <> "\" Then
										filePaths &= "\"
									End If
									Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
									Dim pathLogo As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Societa\1_1.kgb"
									If File.Exists(pathLogo) Then
										Dim pathLogoConv As String = filePaths & "Appoggio\" & Esten & ".jpg"
										Dim c As New CriptaFiles
										c.DecryptFile(CryptPasswordString, pathLogo, pathLogoConv)

										Dim urlLogo As String = mmPaths(2) & "Appoggio/" & Esten & ".jpg"
										filetto = filetto.Replace("***LOGO SOCIETA***", urlLogo)
									Else
										filetto = filetto.Replace("***LOGO SOCIETA***", "")
									End If

									Dim nomeFileHtml As String = filePaths & "Appoggio\" & Esten & ".html"
									Dim nomeFilePDF As String = filePaths & "Appoggio\" & Esten & ".pdf"

									gf.CreaAggiornaFile(nomeFileHtml, filetto)

									Dim pp2 As New pdfGest
									Ritorno = pp2.ConverteHTMLInPDF(nomeFileHtml, nomeFilePDF, "")
									If Ritorno = "*" Then
										Ritorno = "Appoggio/" & Esten & ".pdf"
									End If
								End If
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
End Class