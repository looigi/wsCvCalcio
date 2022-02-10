Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_stat_allti.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsStatAllenamenti
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaStatAllenamentiCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String, Mese As String, NomeSquadra As String, Stampa As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim sMese As String = "/"

				Select Case Mese
					Case "Gennaio"
						sMese = "/01/"
					Case "Febbraio"
						sMese = "/02/"
					Case "Marzo"
						sMese = "/03/"
					Case "Aprile"
						sMese = "/04/"
					Case "Maggio"
						sMese = "/05/"
					Case "Giugno"
						sMese = "/06/"
					Case "Luglio"
						sMese = "/07/"
					Case "Agosto"
						sMese = "/08/"
					Case "Settembre"
						sMese = "/09/"
					Case "Ottobre"
						sMese = "/10/"
					Case "Novembre"
						sMese = "/11/"
					Case "Dicembre"
						sMese = "/12/"
				End Select

				Dim Altro As String = ""
				If TipoDB = "SQLSERVER" Then
					Altro = "And CharIndex(CONVERT(varchar(5),Allenamenti.idCategoria) + '-', Giocatori.Categorie) > 0 "
				Else
					Altro = "And Instr(CONVERT(Allenamenti.idCategoria, varchar(5)) + '-', Giocatori.Categorie) > 0 "
				End If
				Dim Altro2 As String = "Allenamenti.idCategoria=" & idCategoria & " And "
				Dim Altro3 As String = "And idCategoria=" & idCategoria & " "

				If idCategoria = "-1" Then
					Altro = ""
					Altro2 = ""
					Altro3 = ""
				End If
				Try
					Sql = "Select B.idGiocatore, B.Cognome, B.Nome, B.Descrizione,  B.Presenze, B.Totale, (Cast(B.Presenze As Numeric) / Cast(B.Totale As Numeric)) * 100 As Perc, B.NumeroMaglia From ( " &
						"Select A.idGiocatore, A.Cognome, A.Nome, A.Descrizione,  A.Presenze, (SELECT " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " From Allenamenti " &
						"Where idAnno=" & idAnno & " " & Altro3 & " And " & IIf(TipoDB = "SQLSERVER", "CharIndex('" & sMese & "', Datella)>0", "Instr(Datella, '" & sMese & "')>0") & " And Progressivo=0) As Totale, A.NumeroMaglia From ( " &
						"SELECT Giocatori.idGiocatore, Cognome, Nome, Ruoli.Descrizione,  " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " As Presenze, Giocatori.NumeroMaglia " &
						"FROM Allenamenti LEFT JOIN Giocatori ON Allenamenti.idAnno = Giocatori.idAnno AND Allenamenti.idGiocatore=Giocatori.idGiocatore " & Altro & " " &
						"LEFT Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
						"WHERE " & Altro2 & " Allenamenti.idAnno=" & idAnno & " And Giocatori.idGiocatore Is Not Null And " & IIf(TipoDB = "SQLSERVER", "CharIndex('" & sMese & "', Datella)>0", "Instr(Datella, '" & sMese & "', )>0") & " " &
						"Group By Giocatori.idGiocatore, Cognome, Nome, Ruoli.Descrizione, Giocatori.NumeroMaglia " &
						") A) B " &
						"Order By 2"

					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessuna statistica di allenamento rilevata"
						Else
							If Stampa = "NO" Then
								Ritorno = ""
							Else
								Ritorno = "<table style=""width: 100%;"">"
								Ritorno &= "<tr><th>Nominativo</th><th>Ruolo</th><th>Presenze</th><th>Totale</th><th>Perc</th></tr>"
							End If

							Dim q As Integer = 0

							Do Until Rec.Eof()
								Dim perc As String = CInt(Rec("Perc").Value.ToString.Trim)

								If Stampa = "NO" Then
									Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
										Rec("Cognome").Value.ToString.Trim & ";" &
										Rec("Nome").Value.ToString.Trim & ";" &
										Rec("Descrizione").Value.ToString.Trim & ";" &
										Rec("Presenze").Value.ToString.Trim & ";" &
										Rec("Totale").Value.ToString.Trim & ";" &
										perc & ";" &
										Rec("NumeroMaglia").Value.ToString.Trim & ";" &
										"§"
								Else
									q += 1
									Dim conta As String = Format(q, "00")
									Ritorno &= "<tr>"
									Ritorno &= "<td style=""padding-left: 50px;"">" & conta & " - " & Rec("Cognome").Value & " " & Rec("Nome").Value & "</td>"
									Ritorno &= "<td>" & Rec("Descrizione").Value & "</td>"
									Ritorno &= "<td style=""text-align: right;"">" & Rec("Presenze").Value & "</td>"
									Ritorno &= "<td style=""text-align: right;"">" & Rec("Totale").Value & "</td>"
									Ritorno &= "<td style=""text-align: right;"">" & perc & "</td>"
									Ritorno &= "</tr>"
								End If

								Rec.MoveNext()
							Loop
							If Stampa <> "NO" Then
								Ritorno &= "</table>"

								Dim gf As New GestioneFilesDirectory
								Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_allenamenti.txt")

								filetto = filetto.Replace("***TITOLO***", "Allenamenti")
								filetto = filetto.Replace("***DATI***", Ritorno)
								filetto = filetto.Replace("***NOME SQUADRA***", NomeSquadra)

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
								Dim pathLogo As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Societa\1_1.kgb"
								Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
								Dim pathLogoConv As String = filePaths & "Appoggio\" & Esten & ".jpg"
								Dim c As New CriptaFiles
								c.DecryptFile(CryptPasswordString, pathLogo, pathLogoConv)

								Dim urlLogo As String = mmPaths(2) & "Appoggio/" & Esten & ".jpg"
								filetto = filetto.Replace("***LOGO SOCIETA***", urlLogo)

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
	Public Function RitornaInfo(Squadra As String, ByVal idAnno As String, idCategoria As String, idGiocatore As String, Mese As String, NomeSquadra As String, Stampa As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec as object
				Dim Sql As String = ""
				Dim sMese As String = "/"

				Select Case Mese
					Case "Gennaio"
						sMese = "/01/"
					Case "Febbraio"
						sMese = "/02/"
					Case "Marzo"
						sMese = "/03/"
					Case "Aprile"
						sMese = "/04/"
					Case "Maggio"
						sMese = "/05/"
					Case "Giugno"
						sMese = "/06/"
					Case "Luglio"
						sMese = "/07/"
					Case "Agosto"
						sMese = "/08/"
					Case "Settembre"
						sMese = "/09/"
					Case "Ottobre"
						sMese = "/10/"
					Case "Novembre"
						sMese = "/11/"
					Case "Dicembre"
						sMese = "/12/"
				End Select

				Try
					Dim Giocatore As String = ""
					Dim Ok As Boolean = True

					Sql = "SELECT * From Giocatori Where idGiocatore=" & idGiocatore
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
							Ok = False
						Else
							Giocatore = Rec("Cognome").Value & " " & Rec("Nome").Value
						End If
						Rec.Close()
					End If

					If Ok Then
						Sql = "SELECT Allenamenti.Datella, Allenamenti.Orella " &
							"FROM Allenamenti " &
							"WHERE Allenamenti.idAnno=" & idAnno & " AND Allenamenti.idCategoria=" & idCategoria & " AND Allenamenti.idGiocatore=" & idGiocatore & " And " & IIf(TipoDB = "SQLSERVER", "CharIndex('" & sMese & "', Datella)>0", "Instr(Datella,'" & sMese & "')>0") & " " &
							"Order By Datella, Orella"

						Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ritorno = StringaErrore & " Nessuna info di allenamento rilevata"
							Else
								If Stampa = "NO" Then
									Ritorno = ""
								Else
									Ritorno = "<table style=""width: 100%;"">"
									Ritorno &= "<tr><th>Data</th><th>Ora</th></tr>"
								End If
								Do Until Rec.Eof()
									If Stampa = "NO" Then
										Ritorno &= Rec("Datella").Value.ToString & ";" &
										Rec("Orella").Value.ToString.Trim & ";" &
										"§"
									Else
										Ritorno &= "<tr>"
										Ritorno &= "<td style=""margin-left: 50px;"">" & Rec("Datella").Value & "</td><td>" & Rec("Orella").Value & "</td>"
										Ritorno &= "</tr>"
									End If

									Rec.MoveNext()
								Loop
							End If
							Rec.Close()

							If Stampa <> "NO" Then
								Ritorno &= "</table>"

								Dim gf As New GestioneFilesDirectory
								Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_allenamenti.txt")

								filetto = filetto.Replace("***TITOLO***", "Dettaglio allenamenti giocatore")
								filetto = filetto.Replace("***DATI***", Ritorno)
								filetto = filetto.Replace("***NOME SQUADRA***", Giocatore)

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
								Dim pathLogo As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Societa\1_1.kgb"
								Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
								Dim pathLogoConv As String = filePaths & "Appoggio\" & Esten & ".jpg"
								Dim c As New CriptaFiles
								c.DecryptFile(CryptPasswordString, pathLogo, pathLogoConv)

								Dim urlLogo As String = mmPaths(2) & "Appoggio/" & Esten & ".jpg"
								filetto = filetto.Replace("***LOGO SOCIETA***", urlLogo)

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
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

End Class