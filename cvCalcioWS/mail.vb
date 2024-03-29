﻿Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Timers

Public Class mail
	Private pathMail As String = ""

	Public Function SendEmail(pm As String, Squadra As String, Mittente As String, ByVal oggetto As String, ByVal newBody As String, ByVal ricevente As String, ByVal Allegato() As String,
							  Optional AllegatoOMultimedia As String = "", Optional NuovaSocieta As String = "") As String
		Dim Ritorno As String = "*"
		Dim s As New strutturaMail
		s.Squadra = Squadra
		s.Mittente = Mittente
		s.Oggetto = oggetto
		s.newBody = newBody
		s.Ricevente = ricevente
		s.Allegato = Allegato
		s.AllegatoOMultimedia = AllegatoOMultimedia
		s.NuovaSocieta = NuovaSocieta
		s.MP = pm

		' pathMail = HttpContext.Current.Server.MapPath(".")
		pathMail = pm

		listaMails.Add(s)

		If effettuaLog Then
			Dim gf As New GestioneFilesDirectory
			Dim paths As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\PathAllegati.txt")
			Dim pp() As String = paths.Split(";")
			pp(1) = pp(1).Replace(vbCrLf, "")
			If Strings.Right(pp(1), 1) <> "\" Then
				pp(1) = pp(1) & "\"
			End If
			gf.CreaDirectoryDaPercorso(pp(1))
			nomeFileLogMail = pp(1) & "\" & Squadra & "\logMail_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
			gf.CreaDirectoryDaPercorso(nomeFileLogMail)
			'Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
			'Dim Allegati As String = ""
			'For Each a As String In s.Allegato
			'	Allegati &= a & ";"
			'Next
			'gf.ApreFileDiTestoPerScrittura(nomeFileLogMail)
			'gf.ScriveTestoSuFileAperto(Datella & " - Nuova Mail: " & s.Squadra & "/" & s.Mittente & "/" & s.Oggetto & "/" & s.Ricevente & "/" & Allegati & "/" & s.AllegatoOMultimedia)
			'gf.ChiudeFileDiTestoDopoScrittura()
		End If

		avviaTimer()

		Return Ritorno
	End Function

	Private Sub avviaTimer()
		If timerMails Is Nothing Then
			timerMails = New Timer(5000)
			AddHandler timerMails.Elapsed, New ElapsedEventHandler(AddressOf scodaMessaggi)
			timerMails.Start()

			If effettuaLog And nomeFileLogMail <> "" Then
				Dim gf As New GestioneFilesDirectory
				Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

				gf.ApreFileDiTestoPerScrittura(nomeFileLogMail)
				gf.ScriveTestoSuFileAperto(Datella & " - Timer avviato. Mail da scodare: " & listaMails.Count)
				gf.ChiudeFileDiTestoDopoScrittura()
			End If
		End If
	End Sub

	Private Sub scodaMessaggi()
		timerMails.Enabled = False
		Dim mail As strutturaMail = listaMails.Item(0)

		Dim gf As New GestioneFilesDirectory
		If effettuaLog Then
			Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

			gf.ApreFileDiTestoPerScrittura(nomeFileLogMail)
			gf.ScriveTestoSuFileAperto(Datella & " - Scodo Mail: " & mail.Squadra & "/" & mail.Mittente & "/" & mail.Oggetto & "/" & mail.Ricevente)
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Dim Ritorno As String = SendEmailAsincrona(mail.MP, mail.Squadra, mail.Mittente, mail.Oggetto, mail.newBody, mail.Ricevente, mail.Allegato, mail.AllegatoOMultimedia, mail.NuovaSocieta, gf)
		listaMails.RemoveAt(0)
		If listaMails.Count > 0 Then
			timerMails.Enabled = True
		Else
			timerMails = Nothing
			listaMails = New List(Of strutturaMail)
		End If
	End Sub

	Private Function SendEmailAsincrona(Mp As String, Squadra As String, Mittente As String, ByVal oggetto As String, ByVal newBody As String,
										ByVal ricevente As String, ByVal Allegato() As String, AllegatoOMultimedia As String, NuovaSocieta As String,
										gf As GestioneFilesDirectory) As String
		'Dim myStream As StreamReader = New StreamReader(Server.MapPath(ConfigurationManager.AppSettings("VirtualDir") & "mailresponsive.html"))
		'Dim newBody As String = ""
		'newBody = myStream.ReadToEnd()
		'newBody = newBody.Replace("$messaggioemail", body)
		'myStream.Close()
		'myStream.Dispose()

		'CODICE INCOLLATO
		'Dim m As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
		'm.Subject = "subject"
		'm.[To].Add(New System.Net.Mail.MailAddress(Properties.Settings.[Default].RFQRecipient))
		'm.From = New System.Net.Mail.MailAddress(Properties.Settings.[Default].SmtpUsername)
		'm.Body = "message"

		'Try
		'	If fuAttach.HasFile Then m.Attachments.Add(New System.Net.Mail.Attachment(fuAttach.FileContent, fuAttach.FileName))
		'	Dim s As System.Net.Mail.SmtpClient = New System.Net.Mail.SmtpClient(Properties.Settings.[Default].SmtpServer)
		'	s.UseDefaultCredentials = False
		'	s.Credentials = New System.Net.NetworkCredential(Properties.Settings.[Default].SmtpUsername, Properties.Settings.[Default].SmtpPassword)
		'	s.Send(m)
		'Catch ex As Exception
		'	lblError.InnerText = ex.Message
		'End Try
		'End Sub
		'CODICE INCOLLATO


		Dim Ritorno As String = ""
		Dim mail As MailMessage = New MailMessage()
		Dim Credenziali As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\CredenzialiPosta.txt")
		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

		If effettuaLog Then
			gf.ApreFileDiTestoPerScrittura(nomeFileLogMail)
			gf.ScriveTestoSuFileAperto(Datella & " - Inizio")
		End If

		Try
			Dim cr() As String = Credenziali.Split(";")
			Dim Utenza As String = cr(0)
			Dim Password As String = cr(1).Replace(vbCrLf, "")

			If effettuaLog Then
				gf.ScriveTestoSuFileAperto(Datella & " - Inizio 1")
			End If

			If Mittente = "" Then
				Mittente = Utenza
			End If
			'Mittente = Utenza

			If NuovaSocieta <> "" Then
				Dim urlSito As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\PercorsoSito.txt")
				If urlSito.EndsWith("/") Then
					urlSito = Mid(urlSito, 1, urlSito.Length - 1)
				End If

				'Dim contentIDBee As String = "ImageBee"
				'Dim inlineBee = New Attachment(urlSito & "/Scheletri/template_nuova_societa/images/bee.png")
				'inlineBee.ContentId = contentIDBee
				'inlineBee.ContentDisposition.Inline = True
				'inlineBee.ContentDisposition.DispositionType = DispositionTypeNames.Inline

				'mail.Attachments.Add(inlineBee)

				newBody = newBody.Replace("***contentBee***", urlSito & "/Scheletri/template_nuova_societa/images/bee.png")

				'Dim contentIDFB As String = "ImageFB"
				'Dim inlineFB = New Attachment(urlSito & "/Scheletri/template_nuova_societa/images/facebook2x.png")
				'inlineFB.ContentId = contentIDFB
				'inlineFB.ContentDisposition.Inline = True
				'inlineFB.ContentDisposition.DispositionType = DispositionTypeNames.Inline

				'mail.Attachments.Add(inlineFB)

				newBody = newBody.Replace("***contentFB***", urlSito & "/Scheletri/template_nuova_societa/images/facebook2x.png")

				'Dim contentIDLogo As String = "ImageLogo"
				'Dim inlineLogo = New Attachment(urlSito & "/Scheletri/template_nuova_societa/images/LOGOinCalcio200n.png")
				'inlineLogo.ContentId = contentIDLogo
				'inlineLogo.ContentDisposition.Inline = True
				'inlineLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline

				'mail.Attachments.Add(inlineLogo)

				newBody = newBody.Replace("***contentLOGO***", urlSito & "/Scheletri/template_nuova_societa/images/LOGOinCalcio200n.png")

				'Dim contentIDPC As String = "ImagePC"
				'Dim inlinePC = New Attachment(urlSito & "/Scheletri/template_nuova_societa/images/Portatile_homeapp_1.png")
				'inlinePC.ContentId = contentIDPC
				'inlinePC.ContentDisposition.Inline = True
				'inlinePC.ContentDisposition.DispositionType = DispositionTypeNames.Inline

				'mail.Attachments.Add(inlinePC)

				newBody = newBody.Replace("***contentPC***", urlSito & "/Scheletri/template_nuova_societa/images/Portatile_homeapp_1.png")
			End If

			mail.From = New MailAddress(Mittente)
			mail.[To].Add(New MailAddress(ricevente))
			' mail.CC.Add(New MailAddress("email"))
			mail.Subject = oggetto
			mail.IsBodyHtml = True
			If newBody <> "" Then
				mail.Body = newBody ' CreaCorpoMail(Squadra, mail, newBody)
			Else
				mail.Body = ""
			End If

			If effettuaLog Then
				gf.ScriveTestoSuFileAperto(Datella & " - Inizio 2")
			End If

			mail.Body &= "<br><hr />"
			mail.Body &= "<span style=""font-family: Verdana; font-size: 12px;"">Mail inviata tramite InCalcio, software per la gestione delle societa' di calcio - <a href=""www.incalcio.it"">www.incalcio.it</a> - <a href=""mailto:info@incalcio.it"">info@incalcio.it</a></span>"

			Dim Data As Attachment = Nothing
			If Allegato.Length > 0 Then
				For Each All As String In Allegato
					If All <> "" Then
						Dim Allegatone As String = All
						Dim paths As String = ""
						If AllegatoOMultimedia = "A" Then
							'paths = gf.LeggeFileIntero(pathMail & "\Impostazioni\PathAllegati.txt")
							'Dim p() As String = paths.Split(";")
							'If Strings.Right(p(0), 1) <> "\" Then
							'	p(0) &= "\"
							'End If
							'Allegatone = p(0) & Allegatone
							Allegatone = Mp & "\" & Allegatone
							If TipoDB <> "SQLSERVER" Then
								Allegatone = ConvertePath(Allegatone)
							End If
						Else
							If AllegatoOMultimedia = "M" Then
								paths = gf.LeggeFileIntero(pathMail & "\Impostazioni\Paths.txt")
								paths = paths.Replace(vbCrLf, "")
								If Strings.Right(paths, 1) <> "\" Then
									paths &= "\"
								End If
								Allegatone = paths & Allegatone
							End If
						End If

						If effettuaLog Then
							gf.ScriveTestoSuFileAperto(Datella & " - Aggiunge Allegato: " & Allegatone)
						End If

						Data = New Attachment(Allegatone, MediaTypeNames.Application.Octet)
						Dim disposition As ContentDisposition = Data.ContentDisposition
						disposition.CreationDate = System.IO.File.GetCreationTime(Allegatone)
						disposition.ModificationDate = System.IO.File.GetLastWriteTime(Allegatone)
						disposition.ReadDate = System.IO.File.GetLastAccessTime(Allegatone)
						mail.Attachments.Add(Data)
					End If

					If effettuaLog Then
						gf.ScriveTestoSuFileAperto(Datella & " - Inizio 2-1")
					End If
				Next
			End If

			If effettuaLog Then
				gf.ScriveTestoSuFileAperto(Datella & " - Inizio 3")
			End If
			'mail.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8")
			'Dim plainView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(System.Text.RegularExpressions.Regex.Replace(newBody, "< (.|\n) *?>", String.Empty), Nothing, "text/plain")
			'Dim htmlView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(newBody, Nothing, "text/html")
			'mail.AlternateViews.Add(plainView)
			'mail.AlternateViews.Add(htmlView)

			Dim smtpClient As SmtpClient = New SmtpClient("smtps.aruba.it")
			smtpClient.EnableSsl = True
			smtpClient.Port = 587
			smtpClient.UseDefaultCredentials = False
			smtpClient.Credentials = New System.Net.NetworkCredential(Utenza, Password)
			smtpClient.Send(mail)
			smtpClient = Nothing

			'Dim s As System.Net.Mail.SmtpClient = New System.Net.Mail.SmtpClient("smtps.aruba.it")
			's.UseDefaultCredentials = False
			's.Credentials = New System.Net.NetworkCredential(Utenza, Password)
			's.Send(mail)

			If effettuaLog Then
				gf.ScriveTestoSuFileAperto(Datella & " - Invio in corso")
			End If

			If Allegato.Length > 0 And Not Data Is Nothing Then
				Try
					Data.Dispose()
				Catch ex As Exception

				End Try
			End If

			Ritorno = "*"
			If effettuaLog Then
				gf.ScriveTestoSuFileAperto(Datella & " - Invio effettuato")
			End If
		Catch ex As Exception
			Ritorno = StringaErrore & ex.Message

			If effettuaLog Then
				gf.ScriveTestoSuFileAperto(Datella & " - Errore nell'invio: " & ex.Message)
			End If
		End Try
		'smtpClient.Dispose()

		If effettuaLog Then
			gf.ScriveTestoSuFileAperto(Datella & "-----------------------------------------------------------------")
			gf.ScriveTestoSuFileAperto(Datella & "")
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Return Ritorno
	End Function

	Private Function CreaCorpoMail(Squadra As String, mail As MailMessage, newBody As String) As String
		Try
			Dim gf As New GestioneFilesDirectory
			Dim Righe As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\Paths.txt")
			Righe = Righe.Replace(vbCrLf, "")

			Dim Body As String = ""
			'Dim logoApplicazione As String = Righe & "logoApplicazione.png"
			'Dim sfondoMail As String = Righe & "bg.jpg"

			Dim filePaths As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\PathAllegati.txt")
			Dim p() As String = filePaths.Split(";")
			If Strings.Right(p(0), 1) <> "\" Then
				p(0) &= "\"
			End If
			Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_mail.txt"
			If Not ControllaEsistenzaFile(pathFilePosta) Then
				pathFilePosta = pathMail & "\Scheletri\base_mail.txt"
			End If
			Dim Corpo As String = gf.LeggeFileIntero(pathFilePosta)
			'Corpo = Corpo.Replace("***SFONDO***", sfondoMail)
			'Corpo = Corpo.Replace("***LOGO APPLICAZIONE***", logoApplicazione)

			Corpo = Corpo.Replace("***BODY***", "<span style=""font-family: Verdana; font-size: 18px;"">" & newBody & "</span>")

			'Dim contentID As String = "Image1" ' & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
			'Dim inlineLogo = New Attachment(sfondoMail)
			'inlineLogo.ContentId = contentID
			'inlineLogo.ContentDisposition.Inline = True
			'inlineLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline

			'mail.Attachments.Add(inlineLogo)

			'Dim contentID2 As String = "Image2" ' & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
			'Dim inlineLogo2 = New Attachment(logoApplicazione)
			'inlineLogo2.ContentId = contentID2
			'inlineLogo2.ContentDisposition.Inline = True
			'inlineLogo2.ContentDisposition.DispositionType = DispositionTypeNames.Inline

			'mail.Attachments.Add(inlineLogo2)

			'Corpo = Corpo.Replace("***SFONDO***", "cid:" & contentID)
			'Corpo = Corpo.Replace("***LOGO APPLICAZIONE***", "cid:" & contentID2)
			Corpo = Corpo.Replace("***LOGO APPLICAZIONE***", "")

			'gf.ApreFileDiTestoPerScrittura(HttpContext.Current.Server.MapPath(".") & "\MAIL.txt")
			'gf.ScriveTestoSuFileAperto(Corpo)
			'gf.ChiudeFileDiTestoDopoScrittura()
			Return Corpo
		Catch ex As Exception
			Return StringaErrore & ex.Message
		End Try
	End Function

End Class
