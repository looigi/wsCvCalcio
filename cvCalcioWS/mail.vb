Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Timers

Public Class mail
	Public Function SendEmail(Squadra As String, Mittente As String, ByVal oggetto As String, ByVal newBody As String, ByVal ricevente As String, ByVal Allegato As String) As String
		Dim Ritorno As String = "*"
		Dim s As New strutturaMail
		s.Squadra = Squadra
		s.Mittente = Mittente
		s.Oggetto = oggetto
		s.newBody = newBody
		s.Ricevente = ricevente
		s.Allegato = Allegato

		pathMail = HttpContext.Current.Server.MapPath(".")

		listaMails.Add(s)

		avviaTimer()

		Return Ritorno
	End Function

	Private Sub avviaTimer()
		If timerMails Is Nothing Then
			timerMails = New Timer(5000)
			AddHandler timerMails.Elapsed, New ElapsedEventHandler(AddressOf scodaMessaggi)
			timerMails.Start()
		End If
	End Sub

	Private Sub scodaMessaggi()
		timerMails.Enabled = False
		Dim mail As strutturaMail = listaMails.Item(0)
		Dim Ritorno As String = SendEmailAsincrona(mail.Squadra, mail.Mittente, mail.Oggetto, mail.newBody, mail.Ricevente, mail.Allegato)
		listaMails.RemoveAt(0)
		If listaMails.Count > 0 Then
			timerMails.Enabled = True
		Else
			timerMails = Nothing
			listaMails = New List(Of strutturaMail)
		End If
	End Sub

	Private Function SendEmailAsincrona(Squadra As String, Mittente As String, ByVal oggetto As String, ByVal newBody As String, ByVal ricevente As String, ByVal Allegato As String) As String
		'Dim myStream As StreamReader = New StreamReader(Server.MapPath(ConfigurationManager.AppSettings("VirtualDir") & "mailresponsive.html"))
		'Dim newBody As String = ""
		'newBody = myStream.ReadToEnd()
		'newBody = newBody.Replace("$messaggioemail", body)
		'myStream.Close()
		'myStream.Dispose()

		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim mail As MailMessage = New MailMessage()
		Dim Credenziali As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\CredenzialiPosta.txt")

		Try
			Dim cr() As String = Credenziali.Split(";")
			Dim Utenza As String = cr(0)
			Dim Password As String = cr(1).Replace(vbCrLf, "")

			If Mittente = "" Then
				Mittente = Utenza
			End If
			'Mittente = Utenza

			mail.From = New MailAddress(Mittente)
			mail.[To].Add(New MailAddress(ricevente))
			' mail.CC.Add(New MailAddress("email"))
			mail.Subject = oggetto
			mail.IsBodyHtml = True
			mail.Body = CreaCorpoMail(Squadra, mail, newBody)

			Dim Data As Attachment = Nothing
			If Allegato <> "" Then
				Data = New Attachment(Allegato, MediaTypeNames.Application.Octet)
				Dim disposition As ContentDisposition = Data.ContentDisposition
				disposition.CreationDate = System.IO.File.GetCreationTime(Allegato)
				disposition.ModificationDate = System.IO.File.GetLastWriteTime(Allegato)
				disposition.ReadDate = System.IO.File.GetLastAccessTime(Allegato)
				mail.Attachments.Add(Data)
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

			If Allegato <> "" Then
				Data.Dispose()
			End If

			Ritorno = "*"
		Catch ex As Exception
			Ritorno = StringaErrore & ex.Message
		End Try
		'smtpClient.Dispose()

		Return Ritorno
	End Function

	Private Function CreaCorpoMail(Squadra As String, mail As MailMessage, newBody As String) As String
		Try
			Dim gf As New GestioneFilesDirectory
			Dim Righe As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\Paths.txt")
			Righe = Righe.Replace(vbCrLf, "")

			Dim Body As String = ""
			Dim logoApplicazione As String = Righe & "logoApplicazione.png"
			'Dim sfondoMail As String = Righe & "bg.jpg"

			Dim filePaths As String = gf.LeggeFileIntero(pathMail & "\Impostazioni\PathAllegati.txt")
			Dim p() As String = filePaths.Split(";")
			If Strings.Right(p(0), 1) <> "\" Then
				p(0) &= "\"
			End If
			Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_mail.txt"
			If Not File.Exists(pathFilePosta) Then
				pathFilePosta = pathMail & "\Scheletri\base_mail.txt"
			End If
			Dim Corpo As String = gf.LeggeFileIntero(pathFilePosta)
			'Corpo = Corpo.Replace("***SFONDO***", sfondoMail)
			'Corpo = Corpo.Replace("***LOGO APPLICAZIONE***", logoApplicazione)

			Corpo = Corpo.Replace("***BODY***", "<span style=""font-family: Verdana; font-size: 18px;"">" & newBody & "</span>")

			Dim contentID As String = "Image1" ' & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
			'Dim inlineLogo = New Attachment(sfondoMail)
			'inlineLogo.ContentId = contentID
			'inlineLogo.ContentDisposition.Inline = True
			'inlineLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline

			'mail.Attachments.Add(inlineLogo)

			Dim contentID2 As String = "Image2" ' & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
			Dim inlineLogo2 = New Attachment(logoApplicazione)
			inlineLogo2.ContentId = contentID2
			inlineLogo2.ContentDisposition.Inline = True
			inlineLogo2.ContentDisposition.DispositionType = DispositionTypeNames.Inline

			mail.Attachments.Add(inlineLogo2)

			Corpo = Corpo.Replace("***SFONDO***", "cid:" & contentID)
			Corpo = Corpo.Replace("***LOGO APPLICAZIONE***", "cid:" & contentID2)

			'gf.ApreFileDiTestoPerScrittura(HttpContext.Current.Server.MapPath(".") & "\MAIL.txt")
			'gf.ScriveTestoSuFileAperto(Corpo)
			'gf.ChiudeFileDiTestoDopoScrittura()
			Return Corpo
		Catch ex As Exception
			Return StringaErrore & ex.Message
		End Try
	End Function
End Class
