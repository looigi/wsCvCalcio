Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Data.Common
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://cvCalcio.reports.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsReports
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaAnni(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB
			Dim Rec As Object

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = "Select Distinct YEAR(" & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, DataDiNascita, 121)", "Convert(DataDiNascita, DateTime)") & ") As Anno From Giocatori Where Eliminato='N' Order By 1"
				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						Do Until Rec.Eof()
							Ritorno &= Rec("Anno").Value & "§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close()
				End If

			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StampaPagamenti(Squadra As String, Modalita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB
			Dim Rec As Object
			Dim Rec2 As Object

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim Altro As String = ""

				Select Case Modalita
					Case "1"
						Altro = "Where TipoPagamento = 'Rata' "
					Case "2"
						Altro = "Where TipoPagamento = 'Altro' "
					Case "3"
						Altro = ""
					Case "4"
						Altro = "Where Cosa = 'Validato N'"
				End Select

				Sql = "Select * From  ( " &
					"Select *, TotaleDaPagare - TotalePagato As Differenza From ( " &
					"Select 'Validato S' As Cosa, A.idGiocatore, Cognome, Nome, C.Descrizione As DescrizioneQuota, C.Importo, 0) As Importo, " & If(TipoDB = "SQLSERVER", "IsNull(Sconto, 0)", "COALESCE(Sconto, 0)") & " As Sconto, " &
					" " & IIf(TipoDB = "SQLSERVER", "IsNull(C.Importo, 0)", "COALESCE(C.Importo, 0)") & " - " & IIf(TipoDB = "SQLSERVER", "IsNull(Sconto, 0)", "COALESCE(Sconto, 0)") & " As TotaleDaPagare,  " &
					"Validato, Pagamento, idTipoPagamento, " &
					"CASE idTipoPagamento " &
					"     WHEN 1 THEN 'Rata' " &
					"     ELSE 'Altro' " &
					"END As TipoPagamento , A.NumeroRicevuta,  " &
					"(Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Sum(Pagamento),0)", "COALESCE(Sum(Pagamento),0)") & " From GiocatoriPagamenti Where idGiocatore=A.idGiocatore And Eliminato='N' And idTipoPagamento = 1 And Validato='S') As TotalePagato, " &
					"A.DataPagamento, E.MetodoPagamento, A.Note " &
					"From  " &
					"GiocatoriPagamenti A " &
					"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
					"Left Join GiocatoriDettaglio D On D.idGiocatore = A.idGiocatore   " &
					"Left Join Quote C On A.idQuota = C.idQuota " &
					"Left Join MetodiPagamento E On E.idMetodoPagamento = A.MetodoPagamento " &
					"Where A.Eliminato = 'N' And Validato = 'S' " &
					") As A  " &
					"Union All " &
					"Select *, TotaleDaPagare - TotalePagato As Differenza From ( " &
					"Select 'Validato N' As Cosa, A.idGiocatore, Cognome, Nome, C.Descrizione As DescrizioneQuota, " & IIf(TipoDB = "SQLSERVER", "IsNull(C.Importo, 0)", "COALESCE(C.Importo, 0)") & " As Importo, " &
					" " & IIf(TipoDB = "SQLSERVER", "IsNull(Sconto, 0)", "COALESCE(Sconto, 0)") & " As Sconto, " & IIf(TipoDB = "SQLSERVER", "IsNull(C.Importo, 0) - IsNull(Sconto, 0)", "COALESCE(C.Importo, 0) - COALESCE(Sconto, 0)") & " As TotaleDaPagare,  " &
					"Validato, Pagamento, idTipoPagamento, " &
					"CASE idTipoPagamento " &
					"     WHEN 1 THEN 'Rata' " &
					"     ELSE 'Altro' " &
					"END As TipoPagamento , A.NumeroRicevuta,  " &
					"(Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Sum(Pagamento),0)", "COALESCE(Sum(Pagamento),0)") & " From GiocatoriPagamenti Where idGiocatore=A.idGiocatore And Eliminato='N' And idTipoPagamento = 1 And Validato='S') As TotalePagato, " &
					"A.DataPagamento, E.MetodoPagamento, A.Note " &
					"From  " &
					"GiocatoriPagamenti A " &
					"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
					"Left Join GiocatoriDettaglio D On D.idGiocatore = A.idGiocatore   " &
					"Left Join Quote C On A.idQuota = C.idQuota " &
					"Left Join MetodiPagamento E On E.idMetodoPagamento = A.MetodoPagamento " &
					"Where A.Eliminato = 'N' And Validato = 'N' " &
					") As B " &
					") As A  " &
					" " & Altro & " " &
					"Order By Cognome, Nome, Validato"

				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					Dim Output As New StringBuilder
					Dim gf As New GestioneFilesDirectory
					Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					pp = pp.Replace(vbCrLf, "")
					pp = pp.Trim()
					If Strings.Right(pp, 1) = "\" Then
						pp = Mid(pp, 1, pp.Length - 1)
					End If
					Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
					Dim P() As String = PathAllegati.Split(";")
					P(2) = P(2).Replace(vbCrLf, "").Trim
					If Strings.Right(P(2), 1) = "\" Then
						P(2) = Mid(P(2), 1, P(2).Length - 1)
					End If

					Dim CodSquadra() As String = Squadra.Split("_")
					Dim idSquadra As Integer = Val(CodSquadra(1))
					Dim idAnno As String = Val(CodSquadra(0)).ToString.Trim
					Dim AnnoAttivazione As String = ""
					Dim NomeSquadra As String = ""
					Dim c As New CriptaFiles

					Dim totQuota As Single = 0
					Dim totPagatoQuota As Single = 0
					Dim totPagatoManuale As Single = 0
					Dim totDifferenza As Single = 0
					Dim Vecchia As String = ""

					If Modalita = "PDF" Then
						Output.Append("<table style=""width:  100%;"" cellapadding=""0"" cellspacing=""0"">")
						Output.Append("<tr>")
						Output.Append("<th></th>")
						Output.Append("<th style = ""text-align: left;"">Cognome</th>")
						Output.Append("<th style=""text-align: left;"">Nome</th>")
						' Output.Append("<th style = ""text-align: left;"">Importo Quota</th>")
						Output.Append("<th style = ""text-align: right;"">Totale Quota</th>")
						'Output.Append("<th style = ""text-align: right;"">Pag. Automatico</th>")
						'Output.Append("<th style = ""text-align: right;"">Pag. Manuale</th>")
						Output.Append("<th style = ""text-align: right;"">Sconto</th>")
						Output.Append("<th style = ""text-align: right;"">Differenza</th>")
						Output.Append("<th style = ""text-align: right;"">Quota</th>")
						Output.Append("<th style = ""text-align: right;"">Pagamento Automatico</th>")
						'Output.Append("<th style = ""text-align: right;"">Pagamento Manuale</th>")
						Output.Append("<th style = ""text-align: right;"">DataPagamento</th>")
						Output.Append("<th style = ""text-align: right;"">Validato</th>")
						Output.Append("<th style = ""text-align: right;"">MetodoPagamento</th>")
						Output.Append("<th style = ""text-align: right;"">NumeroRicevuta</th>")
						Output.Append("<th style = ""text-align: right;"">Tipo Pagamento</th>")
						Output.Append("<th style = ""text-align: right;"">Note</th>")
						Output.Append("</tr>")
					Else
						Output.Append("Cognome;Nome;Totale Quota;Sconto;Differenza;Quota;Pagamento Totale;DataPagamento;Validato;MetodoPagamento;NumeroRicevuta;Tipo Pagamento;Note;")
						Output.Append(vbCrLf)
					End If

					Do Until Rec.Eof()
						If Modalita = "PDF" Then
							Dim Immagine As String = ""
							Dim Esten2 As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
							Dim urlImmagine As String = P(2) & "/" & NomeSquadra.Replace(" ", "_") & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
							Dim pathImmagine As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Giocatori\" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
							Dim urlImmagineConvertita As String = P(2) & "/Appoggio/" & Rec("idGiocatore").Value & "_" & Esten2 & ".png"
							Dim pathImmagineConvertita As String = pp & "\Appoggio\" & Rec("idGiocatore").Value & "_" & Esten2 & ".png"
							If ControllaEsistenzaFile(pathImmagine) Then
								c.DecryptFile(CryptPasswordString, pathImmagine, pathImmagineConvertita)

								Immagine = "<img src=""" & urlImmagineConvertita & """ style=""width: 50px; height: 50px;"" />"
							Else
								urlImmagineConvertita = P(2) & "/Sconosciuto.png"
								Immagine = "<img src=""" & urlImmagineConvertita & """ style=""width: 50px; height: 50px;"" />"
							End If

							Output.Append("<tr>")
							Output.Append("<td>" & Immagine & "</td>")
							Output.Append("<td>" & Rec("Cognome").Value & "</td>")
							Output.Append("<td>" & Rec("Nome").Value & "</td>")
							' Output.Append("<td>" & Rec("Descrizione").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("Importo").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("ImportoQuota").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("ImportoManuale").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("Sconto").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("Differenza").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("Quota").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("Pagamento").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("PagamentoManuale").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("DataPagamento").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("Validato").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("MetodoPagamento").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("NumeroRicevuta").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("TipoPagamento").Value & "</td>")
							Output.Append("<td style = ""text-align: right;"">" & Rec("Note").Value & "</td>")
							Output.Append("</tr>")
						Else
							'Dim idRate As String = "" & Rec("idRata").Value
							'idRate = idRate.Replace(";", "-")
							'If idRate.EndsWith("-") Then
							'	idRate = Mid(idRate, 1, idRate.Length - 1)
							'End If

							'Cosa    idGiocatore	Cognome	Nome	DescrizioneQuota	Importo	Sconto	TotaleDaPagare	Validato	Pagamento	idTipoPagamento	TipoPagamento	NumeroRicevuta	TotalePagato	Differenza

							Dim Riga As String = "" & Rec("Cognome").Value & ";" & "" & Rec("Nome").Value & ";"
							Dim RigaDaScrivere As String = "" & Rec("Cognome").Value & ";" & "" & Rec("Nome").Value & ";" & "" & Rec("Importo").Value & ";" & "" & Rec("Sconto").Value & ";" & "" & Rec("Differenza").Value & ";" & "" & Rec("DescrizioneQuota").Value & ";"

							If Riga <> Vecchia Then
								Vecchia = Riga
								'Output.Append("" & Rec("Cognome").Value & ";")
								'Output.Append("" & Rec("Nome").Value & ";")
								'' Output.Append(Rec("Descrizione").Value & ";")
								'Output.Append("" & Rec("Importo").Value & ";")
								''Output.Append("" & Rec("ImportoQuota").Value & ";")
								''Output.Append("" & Rec("ImportoManuale").Value & ";")
								'Output.Append("" & Rec("Sconto").Value & ";")
								'Output.Append("" & Rec("Differenza").Value & ";")
								'Output.Append("" & Rec("Quota").Value & ";")
								Output.Append(RigaDaScrivere)
							Else
								Output.Append(";;;;;;")
							End If
							Output.Append("" & Rec("Pagamento").Value & ";")
							'Output.Append("" & Rec("PagamentoManuale").Value & ";")
							Output.Append("" & Rec("DataPagamento").Value & ";")
							Output.Append("" & Rec("Validato").Value & ";")
							Output.Append("" & Rec("MetodoPagamento").Value & ";")
							Output.Append("" & Rec("NumeroRicevuta").Value & ";")
							Output.Append("" & Rec("TipoPagamento").Value & ";")
							Output.Append("" & Rec("Note").Value & ";")
							Output.Append(vbCrLf)
						End If

						totQuota += Val(("" & Rec("Pagamento").Value))
						'totPagatoQuota += Val(("" & Rec("ImportoQuota").Value))
						'totPagatoManuale += Val(("" & Rec("ImportoManuale").Value))
						'totDifferenza += Val(("" & Rec("Differenza").Value))

						Rec.MoveNext()
					Loop
					Rec.Close()

					Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
					filePaths = filePaths.Replace(vbCrLf, "")
					If Strings.Right(filePaths, 1) <> "\" Then
						filePaths &= "\"
					End If
					Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

					If Modalita = "PDF" Then
						Output.Append("<tr>")
						Output.Append("<td colspan=""8""><hr />")
						Output.Append("</td>")
						Output.Append("</tr>")

						Output.Append("<tr>")
						'Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						Output.Append("<td></td>")
						'Output.Append("<td style = ""font-weight: bold; text-align: right;"">" & totQuota & "</td>")
						'Output.Append("<td style = ""font-weight: bold; text-align: right;"">" & totPagatoQuota & "</td>")
						'Output.Append("<td style = ""font-weight: bold; text-align: right;"">" & totPagatoManuale & "</td>")
						'Output.Append("<td style = ""font-weight: bold; text-align: right;""></td>")
						Output.Append("<td style = ""font-weight: bold; text-align: right;"">" & totQuota & "</td>")
						Output.Append("</tr>")

						Output.Append("</table>")

						Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_report.txt")

						filetto = filetto.Replace("***TITOLO***", "Lista Pagamenti")
						filetto = filetto.Replace("***DATI***", Output.ToString)
						filetto = filetto.Replace("***NOME SQUADRA***", "<br /><br />" & NomeSquadra)

						Dim multimediaPaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
						Dim mmPaths() As String = multimediaPaths.Split(";")
						mmPaths(2) = mmPaths(2).Replace(vbCrLf, "")
						If Strings.Right(mmPaths(2), 1) <> "/" Then
							mmPaths(2) &= "/"
						End If

						Dim pathLogo As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Societa\1_1.kgb"
						If ControllaEsistenzaFile(pathLogo) Then
							Dim pathLogoConv As String = filePaths & "Appoggio\" & Esten & ".jpg"
							c.DecryptFile(CryptPasswordString, pathLogo, pathLogoConv)

							Dim urlLogo As String = mmPaths(2) & "Appoggio/" & Esten & ".jpg"
							filetto = filetto.Replace("***LOGO SOCIETA***", urlLogo)
						Else
							filetto = filetto.Replace("***LOGO SOCIETA***", "")
						End If

						' filetto &= "<hr />Stampato tramite InCalcio, software per la gestione delle società di calcio - www.incalcio.it - info@incalcio.it"

						Dim nomeFileHtml As String = filePaths & "Appoggio\" & Esten & ".html"
						Dim nomeFilePDF As String = filePaths & "Appoggio\" & Esten & ".pdf"

						gf.CreaAggiornaFile(nomeFileHtml, filetto)

						Dim pp2 As New pdfGest
						Ritorno = pp2.ConverteHTMLInPDF(nomeFileHtml, nomeFilePDF, "")
						If Ritorno = "*" Then
							Ritorno = "Appoggio/" & Esten & ".pdf"
						End If
					Else
						Output.Append("TOTALI;")
						Output.Append(";")
						Output.Append(";")
						Output.Append(";")
						Output.Append(";")
						Output.Append(";")
						Output.Append(totQuota & ";")
						'Output.Append(totPagatoQuota & ";")
						'Output.Append(totPagatoManuale & ";")
						'Output.Append(totDifferenza & ";")
						Output.Append(vbCrLf)

						Dim nomeFileCSV As String = filePaths & "Appoggio\" & Esten & ".csv"
						gf.CreaAggiornaFile(nomeFileCSV, Output.ToString)

						Ritorno = "Appoggio/" & Esten & ".csv"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StampaQuote(Squadra As String, Anno As String, QuotaPresente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB
			Dim Rec As Object
			Dim Sql As String = ""
			Dim Altro As String = ""

			If Anno <> "" Then
				Altro = "And Year(DataDiNascita)=" & Anno
			End If

			If QuotaPresente = "N" Then
				Sql = "Select Cognome, Nome, DataDiNascita, '' AS Descrizione,  0 As DaPagare, 0 As Sconto, 0 As TotalePagato, 0 As Differenza " &
					"From " &
					"Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Left Join Quote C On B.idQuota = C.idQuota And C.Eliminato = 'N' " &
					"Where A.Eliminato = 'N' And C.Descrizione Is Null " & Altro & " " &
					"Order By Cognome, Nome"
			Else
				Sql = "Select Cognome, Nome, DataDiNascita, Descrizione, DaPagare, Sconto, TotalePagato, (DaPagare - Sconto) - TotalePagato As Differenza From ( " &
					"Select Cognome, Nome, DataDiNascita, C.Descrizione, " & IIf(TipoDB = "SQLSERVER", "IsNull(C.Importo, 0)", "COALESCE(C.Importo, 0)") & " As DaPagare, " & IIf(TipoDB = "SQLSERVER", "IsNull(B.Sconto, 0)", "COALESCE(B.Sconto, 0)") & " As Sconto, " &
					"(Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Sum(Pagamento),0)", "COALESCE(Sum(Pagamento),0)") & " From GiocatoriPagamenti Where idGiocatore=A.idGiocatore And Eliminato='N' And idTipoPagamento = 1 And Validato='S') As TotalePagato " &
					"From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Left Join Quote C On B.idQuota = C.idQuota " &
					"Where A.Eliminato='N' And C.Eliminato='N' " & Altro & " " &
					") As A Order By Cognome, Nome"
			End If
			Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				Dim Output As New StringBuilder
				Dim gf As New GestioneFilesDirectory
				Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
				pp = pp.Replace(vbCrLf, "")
				pp = pp.Trim()
				If Strings.Right(pp, 1) = "\" Then
					pp = Mid(pp, 1, pp.Length - 1)
				End If
				Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim P() As String = PathAllegati.Split(";")
				P(2) = P(2).Replace(vbCrLf, "").Trim
				If Strings.Right(P(2), 1) = "\" Then
					P(2) = Mid(P(2), 1, P(2).Length - 1)
				End If

				Dim CodSquadra() As String = Squadra.Split("_")

				Output.Append("Cognome;Nome;Data Di Nascita;Quota;Da Pagare;Sconto;Pagato;Differenza;")
				Output.Append(vbCrLf)

				Do Until Rec.Eof()
					Output.Append("" & Rec("Cognome").Value & ";")
					Output.Append("" & Rec("Nome").Value & ";")
					Output.Append("" & Rec("DataDiNascita").Value & ";")
					Output.Append("" & Rec("Descrizione").Value & ";")
					Output.Append("" & Rec("DaPagare").Value & ";")
					Output.Append("" & Rec("Sconto").Value & ";")
					Output.Append("" & Rec("TotalePagato").Value & ";")
					Output.Append("" & Rec("Differenza").Value & ";")
					Output.Append(vbCrLf)

					Rec.MoveNext()
				Loop
				Rec.Close()

				Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
				filePaths = filePaths.Replace(vbCrLf, "")
				If Strings.Right(filePaths, 1) <> "\" Then
					filePaths &= "\"
				End If
				Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
				Dim nomeFileCSV As String = filePaths & "Appoggio\" & Esten & ".csv"
				gf.CreaAggiornaFile(nomeFileCSV, Output.ToString)

				Ritorno = "Appoggio/" & Esten & ".csv"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StampaAnagrafica(Squadra As String, Tipologia As String, Dato As String, Certificato As String, FirmePresenti As String, KitConsegnato As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB
			Dim Rec as object
			Dim Rec2 as object

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim CodSquadra() As String = Squadra.Split("_")
				Dim idSquadra As Integer = Val(CodSquadra(1))
				Dim idAnno As String = Val(CodSquadra(0)).ToString.Trim
				Dim AnnoAttivazione As String = ""
				Dim NomeSquadra As String = ""
				Dim IscrFirmaEntrambi As String = ""
				Dim c As New CriptaFiles

				Sql = "Select * From Anni"
				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() = False Then
						IscrFirmaEntrambi = "" & Rec("iscrFirmaEntrambi").Value
					Else
						Ritorno = StringaErrore & " Nessun dato societario rilevato"
						Ok = False
					End If
				End If

				If Ok Then
					Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & idSquadra
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						If Rec.Eof() = False Then
							AnnoAttivazione = "" & Rec("AnnoAttivazione").Value
							NomeSquadra = "" & Rec("Descrizione").Value
						Else
							Ritorno = StringaErrore & " Nessun anno di attivazione rilevato"
							Ok = False
						End If
					End If
				End If

				If Ok Then
					Dim Titolo As String = "Report Anagrafica"
					Dim Altro As String = " "

					Sql = "Select * From Giocatori A " &
						"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						"Where "

					If Tipologia = "1" Then
						' Stampa per anno
						If Dato <> "-1" Then
							Sql &= "YEAR(" & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, DataDiNascita, 121)", "Convert(DataDiNascita, DateTime)") & ") = " & Dato
							Altro &= "per anno"
						Else
							Sql &= " 1=1"
							Altro &= "per tutti gli anni"
						End If
					Else
						If Dato = "1" Then
							' Stampa per scuola calcio
							Sql &= "YEAR(" & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, DataDiNascita ,121)", "Convert(DataDiNascita ,DateTime)") & ") >= Year(CURRENT_TIMESTAMP) - 12"
							Altro &= "per scuola calcio"
						Else
							' Stampa per agonistica
							Sql &= "YEAR(" & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, DataDiNascita ,121)", "Convert(DataDiNascita ,DateTime)") & ") < Year(CURRENT_TIMESTAMP) - 12 And " &
												 "YEAR(" & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, DataDiNascita ,121)", "Convert(DataDiNascita ,DateTime)") & ") >= Year(CURRENT_TIMESTAMP) - 25"
							Altro &= "per agonistica"
						End If
					End If

					Select Case Certificato
						Case "1"
							' Scaduto
							Sql &= " And (B.ScadenzaCertificatoMedico Is Not Null And B.ScadenzaCertificatoMedico <> '' And " & IIf(TipoDB = "SQLSERVER", "DateTime, B.ScadenzaCertificatoMedico ,121)", "Convert(B.ScadenzaCertificatoMedico ,DateTime)") & " <= CURRENT_TIMESTAMP And B.CertificatoMedico = 'S')" ' And (Convert(DateTime, B.ScadenzaCertificatoMedico ,121) > DateAdd(Day, -30, CURRENT_TIMESTAMP) And B.CertificatoMedico = 'S')"
							Altro &= ", certificato medico scaduto"
						Case "2"
							' Presente
							Sql &= " And B.CertificatoMedico = 'S'  And " & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico ,121)", "Convert(B.ScadenzaCertificatoMedico ,DateTime)") & " > CURRENT_TIMESTAMP"
							Altro &= ", certificato medico presente"
						Case "3"
							' Assente
							Sql &= " And (B.CertificatoMedico Is Null Or B.CertificatoMedico = '' Or B.CertificatoMedico = 'N' Or (B.CertificatoMedico = 'S' And (B.ScadenzaCertificatoMedico Is Null Or B.ScadenzaCertificatoMedico = ''))) "
							Altro &= ", certificato medico assente"
						Case "4"
							' In scadenza
							Sql &= " And (" & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico ,121)", "Convert(B.ScadenzaCertificatoMedico ,DateTime)") & " > CURRENT_TIMESTAMP) And " &
								"(" & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico ,121)", "Convert(B.ScadenzaCertificatoMedico ,DateTime)") & " <= " & IIf(TipoDB = "SQLSERVER", "DateAdd(Day, 30, CURRENT_TIMESTAMP)", "ADDDATE(CURRENT_TIMESTAMP, 30)") & " And B.CertificatoMedico = 'S')"
							Altro &= ", certificato medico in scadenza"
					End Select

					Sql &= " And A.Eliminato = 'N' Order By Cognome, Nome"

					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Dim Output As String = ""
						Dim gf As New GestioneFilesDirectory
						Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
						pp = pp.Replace(vbCrLf, "")
						pp = pp.Trim()
						If Strings.Right(pp, 1) = "\" Then
							pp = Mid(pp, 1, pp.Length - 1)
						End If
						Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
						Dim P() As String = PathAllegati.Split(";")
						P(2) = P(2).Replace(vbCrLf, "").Trim
						If Strings.Right(P(2), 1) = "\" Then
							P(2) = Mid(P(2), 1, P(2).Length - 1)
						End If

						Output = "<table style=""width: 100%;"" cellapadding=""0"" cellspacing=""0"">"
						Output &= "<tr><th></th><th style=""text-align: left;"">Cognome</th><th style=""text-align: left;"">Nome</th><th style=""text-align: left;"">Data di nascita</th><th style=""text-align: left;"">Numero Maglia</th><th style=""text-align: left;"">Matricola</th>"
						If Val(Certificato) > 0 Then
							Output &= "<th>Data Scad. Cert.</th>"
						End If
						If Val(KitConsegnato) > 0 Then
							Output &= "<th>Kit</th>"
							Output &= "<th>Taglia</th>"
							Output &= "<th>Dettaglio</th>"
						End If
						Output &= "</tr>"
						Output &= "<tr><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td>"
						If Val(Certificato) > 0 Then
							Output &= "<th><hr /></th>"
						End If
						If Val(KitConsegnato) > 0 Then
							Output &= "<th><hr /></th>"
							Output &= "<th><hr /></th>"
							Output &= "<th><hr /></th>"
						End If
						Output &= "</tr>"

						If FirmePresenti <> "0" Then
							If FirmePresenti = "1" Then
								Altro &= ", firme presenti"
							Else
								Altro &= ", firme non presenti"
							End If
						End If

						If KitConsegnato <> "0" Then
							If KitConsegnato = "1" Then
								Altro &= ", kit consegnato"
							Else
								If KitConsegnato = "2" Then
									Altro &= ", kit non consegnato"
								Else
									Altro &= ", consegnato parzialmente"
								End If
							End If
						End If

						Dim Quanti As Integer = 0

						Do Until Rec.Eof()
							Dim Stampa As Boolean = True

							If FirmePresenti <> "0" Then

								Dim urlFirma As String = ""
								Dim CiSonoFirme As Boolean = True

								' Query di controllo
								'Select Case A.Maggiorenne, GenitoriSeparati, Genitore1, MailGenitore1, Genitore2, MailGenitore2, AffidamentoCongiunto, 
								'                        AbilitaFirmaGenitore1, AbilitaFirmaGenitore2, FirmaAnalogicaGenitore1, FirmaAnalogicaGenitore2 
								'From Giocatori A
								'Left Join GiocatoriDettaglio B On A.idGiocatore=B.idGiocatore
								'Where Cognome Like '%petr%' 

								If "" & Rec("Maggiorenne").Value = "S" Then
									If "" & Rec("AbilitaFirmaGenitore3").Value = "S" Then
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_3.kgb"
										If Not ControllaEsistenzaFile(urlFirma) Then
											CiSonoFirme = False
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore3").Value = "N" Then
											CiSonoFirme = False
										End If
									End If
								Else
									If "" & Rec("GenitoriSeparati").Value = "S" Then
										If "" & Rec("AffidamentoCongiunto").Value = "S" Then
											If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
												urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
												If Not ControllaEsistenzaFile(urlFirma) Then
													CiSonoFirme = False
												End If
											Else
												If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
													CiSonoFirme = False
												End If
											End If

											If Stampa Then
												If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
													urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
													If Not ControllaEsistenzaFile(urlFirma) Then
														CiSonoFirme = False
													End If
												Else
													If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
														CiSonoFirme = False
													End If
												End If
											End If
										Else
											If "" & Rec("idTutore").Value = "1" Then
												If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
													urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
													If Not ControllaEsistenzaFile(urlFirma) Then
														CiSonoFirme = False
													End If
												Else
													If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
														CiSonoFirme = False
													End If
												End If
											Else
												If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
													urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
													If Not ControllaEsistenzaFile(urlFirma) Then
														CiSonoFirme = False
													End If
												Else
													If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
														CiSonoFirme = False
													End If
												End If
											End If
										End If
									Else
										If IscrFirmaEntrambi = "S" Then
											If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
												urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
												If Not ControllaEsistenzaFile(urlFirma) Then
													CiSonoFirme = False
												End If
											Else
												If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
													CiSonoFirme = False
												End If
											End If

											If Stampa Then
												If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
													urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
													If Not ControllaEsistenzaFile(urlFirma) Then
														CiSonoFirme = False
													End If
												Else
													If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
														CiSonoFirme = False
													End If
												End If
											End If
										Else
											If "" & Rec("Genitore1").Value <> "" Then
												If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
													urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
													If Not ControllaEsistenzaFile(urlFirma) Then
														CiSonoFirme = False
													End If
												Else
													If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
														CiSonoFirme = False
													End If
												End If
											Else
												If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
													urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
													If Not ControllaEsistenzaFile(urlFirma) Then
														CiSonoFirme = False
													End If
												Else
													If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
														CiSonoFirme = False
													End If
												End If
											End If
										End If
									End If
								End If

								If FirmePresenti = "1" Then
									If CiSonoFirme = False Then
										Stampa = False
									End If
								Else
									If CiSonoFirme = True Then
										Stampa = False
									End If
								End If
							End If

							Dim NomeKit As String = ""
							Dim TagliaKit As String = ""
							Dim DettaglioKit As String = ""

							If Stampa Then
								If KitConsegnato <> "0" Then
									Sql = "Select C.Quantita, QuantitaConsegnata, D.Descrizione As NomeKit, F.Descrizione As Taglia, G.Descrizione As Elemento From KitComposizione A " &
										"Left Join KitGiocatori B On A.idTipoKit = B.idTipokit And A.idElemento = B.idElemento " &
										"Left Join KitComposizione C On B.idTipoKit = C.idTipoKit And B.idElemento = C.idElemento " &
										"Left Join KitTipologie D On D.idTipoKit = C.idTipoKit " &
										"Left Join Giocatori E On B.idGiocatore = E.idGiocatore " &
										"Left Join Taglie F On E.idTaglia = F.idTaglia " &
										"Left Join KitElementi G On G.idElemento = C.idElemento " &
										"Where B.idGiocatore = " & Rec("idGiocatore").Value & " And C.Eliminato='N' And A.Eliminato='N' And D.Eliminato='N' And E.Eliminato='N' And G.Eliminato='N'"
									Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
										Ok = False
									Else
										Dim Tutto As Boolean = True
										Dim Qualcosa As Boolean = False

										If Rec2.Eof() Then
											Tutto = False
										Else
											Do Until Rec2.Eof()
												If NomeKit = "" Then
													NomeKit = "" & Rec2("NomeKit").Value
													TagliaKit = "" & Rec2("Taglia").Value
												End If

												'If Val(Rec2("QuantitaConsegnata").Value) > 0 Then
												Qualcosa = True
												If Val(Rec2("QuantitaConsegnata").Value) < Val(Rec2("Quantita").Value) Then
													Tutto = False
												End If
												'Else
												'	Tutto = False
												'End If

												'If Val(Rec2("QuantitaConsegnata").Value) > Val(Rec2("Quantita").Value) Then
												'Else
												'Tutto = False
												'End If

												DettaglioKit &= Rec2("Elemento").Value & ": " & Rec2("QuantitaConsegnata").Value & "/" & Rec2("Quantita").Value & "<br />"

												Rec2.MoveNext()
											Loop
											Rec2.Close()
										End If

										' DettaglioKit &= "<br /><br />Tutto: " & Tutto & "<br />Qualcosa: " & Qualcosa

										If KitConsegnato = "1" Then
											' Kit consegnato Sì
											If Tutto = False Then
												Stampa = False
											End If
										End If

										If KitConsegnato = "2" Then
											' Kit consegnato No
											If Tutto = True Or Qualcosa = True Then
												Stampa = False
											End If
										End If

										If KitConsegnato = "3" Then
											' Kit consegnato Parziale
											If Qualcosa = False Or Tutto = True Then
												Stampa = False
											End If
										End If
									End If
								End If
							End If

							If Stampa Then
								Dim ddn As String = "" & Rec("DataDiNascita").Value
								If ddn <> "" Then
									Dim d() As String = ddn.Split("-")
									ddn = d(2) & "/" & d(1) & "/" & d(0)
								End If

								'Dim Immagine As String = ""
								'Dim Esten2 As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
								'Dim urlImmagine As String = P(2) & "/" & NomeSquadra.Replace(" ", "_") & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
								'Dim pathImmagine As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Giocatori\" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
								'Dim urlImmagineConvertita As String = P(2) & "/Appoggio/" & Rec("idGiocatore").Value & "_" & Esten2 & ".png"
								'Dim pathImmagineConvertita As String = pp & "\Appoggio\" & Rec("idGiocatore").Value & "_" & Esten2 & ".png"
								'If ControllaEsistenzaFile(pathImmagine) Then
								'	c.DecryptFile(CryptPasswordString, pathImmagine, pathImmagineConvertita)

								'	Immagine = "<img src=""" & urlImmagineConvertita & """ style=""width: 50px; height: 50px;"" />"
								'Else
								'	urlImmagineConvertita = P(2) & "/Sconosciuto.png"
								'	Immagine = "<img src=""" & urlImmagineConvertita & """ style=""width: 50px; height: 50px;"" />"
								'End If

								Output &= "<tr><td></td><td>" & Rec("Cognome").Value & "</td><td>" & Rec("Nome").Value & "</td><td>" & ddn & "</td><td>" & Rec("NumeroMaglia").Value & "</td><td>" & Rec("Matricola").Value & "</td>"
								If Val(Certificato) > 0 Then
									Dim d As String = "" & Rec("ScadenzaCertificatoMedico").Value
									Dim sData As String = ""
									If d.Contains("-") Then
										Dim dd() As String = d.Split("-")
										sData = dd(2) & "/" & dd(1) & "/" & dd(0)
									End If
									Output &= "<th>" & sData & "</th>"
								End If
								If Val(KitConsegnato) > 0 Then
									Output &= "<th>" & NomeKit & "</th>"
									Output &= "<th>" & TagliaKit & "</th>"
									Output &= "<th style=""text-align: right;"">" & DettaglioKit & "</th>"
								End If
								Output &= "</tr>"
								Output &= "<tr><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td><td><hr /></td>"
								If Val(Certificato) > 0 Then
									Output &= "<th><hr /></th>"
								End If
								If Val(KitConsegnato) > 0 Then
									Output &= "<th><hr /></th>"
									Output &= "<th><hr /></th>"
									Output &= "<th><hr /></th>"
								End If
								Output &= "</tr>"

								Quanti += 1
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()

						Output &= "</table>"

						Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_report.txt")

						filetto = filetto.Replace("***TITOLO***", Titolo & Altro & "<br />Rilevati: " & Quanti)
						filetto = filetto.Replace("***DATI***", Output)
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
						If ControllaEsistenzaFile(pathLogo) Then
							Dim pathLogoConv As String = filePaths & "Appoggio\" & Esten & ".jpg"
							c.DecryptFile(CryptPasswordString, pathLogo, pathLogoConv)

							Dim urlLogo As String = mmPaths(2) & "Appoggio/" & Esten & ".jpg"
							filetto = filetto.Replace("***LOGO SOCIETA***", urlLogo)
						Else
							filetto = filetto.Replace("***LOGO SOCIETA***", "")
						End If

						' filetto &= "<hr />Stampato tramite InCalcio, software per la gestione delle società di calcio - www.incalcio.it - info@incalcio.it"

						Dim nomeFileHtml As String = filePaths & "Appoggio\" & Esten & ".html"
						'Dim nomeFilePDF As String = filePaths & "Appoggio\" & Esten & ".pdf"

						gf.CreaAggiornaFile(nomeFileHtml, filetto)

						'Dim pp2 As New pdfGest
						'Ritorno = pp2.ConverteHTMLInPDF(nomeFileHtml, nomeFilePDF, "")
						'If Ritorno = "*" Then
						Ritorno = "Appoggio/" & Esten & ".html"
						'End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class