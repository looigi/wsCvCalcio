Imports System.IO
Imports System.Net.Mail

Public Class mail
	Public Function SendEmail(ByVal oggetto As String, ByVal newBody As String, ByVal Mittente As String, ByVal Optional ricevente As String = "emaildefault") As String
		'Dim myStream As StreamReader = New StreamReader(Server.MapPath(ConfigurationManager.AppSettings("VirtualDir") & "mailresponsive.html"))
		'Dim newBody As String = ""
		'newBody = myStream.ReadToEnd()
		'newBody = newBody.Replace("$messaggioemail", body)
		'myStream.Close()
		'myStream.Dispose()

		Dim Ritorno As String = ""
		Dim mail As MailMessage = New MailMessage()
		mail.From = New MailAddress(Mittente)
		mail.[To].Add(New MailAddress(ricevente))
		' mail.CC.Add(New MailAddress("email"))
		mail.Subject = oggetto
		mail.Body = newBody
		mail.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8")
		Dim plainView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(System.Text.RegularExpressions.Regex.Replace(newBody, "< (.|\n) *?>", String.Empty), Nothing, "text/plain")
		Dim htmlView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(newBody, Nothing, "text/html")
		mail.AlternateViews.Add(plainView)
		mail.AlternateViews.Add(htmlView)
		mail.IsBodyHtml = True
		Dim smtpClient As SmtpClient = New SmtpClient("smtps.aruba.it")
		smtpClient.EnableSsl = True
		smtpClient.Port = 587
		smtpClient.UseDefaultCredentials = False
		smtpClient.Credentials = New System.Net.NetworkCredential("notifiche@incalcio.cloud", "Ch10d3ll1184!")
		Try
			smtpClient.Send(mail)
			Ritorno = "*"
		Catch ex As Exception
			Ritorno = "ERROR: " & ex.Message
		End Try
		'smtpClient.Dispose()
		smtpClient = Nothing

		Return Ritorno
	End Function
End Class
