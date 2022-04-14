Imports System.IO
Imports OpaqueMail

Public Class mailImap
	Private Function SistemaStringa(Stringa As String) As String
		Dim S As String = Stringa

		If S Is Nothing Then S = ""

		S = S.Replace(";", "***PV***")
		S = S.Replace("|", "***PIPE***")
		S = S.Replace("'", "''")

		Return S
	End Function

	Public Function RitornaMessaggi(MP As String, Squadra As String, idAnno As String, idUtente As String, Casella As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(MP, Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
				Dim Rec2 As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim idMail As Integer = 0
				Dim Ok As Boolean = True
				Dim QuanteMails As Integer = 0
				Dim Utenza As String = ""
				Dim Password As String = ""
				Dim UltimoControllo As String = ""
				Dim gf As New GestioneFilesDirectory
				Dim Righe As String = gf.LeggeFileIntero(MP & "\Impostazioni\PercorsoAttachment.txt")
				Dim pathAttachments = Righe.Replace(vbCrLf, "")
				If pathAttachments.EndsWith("\") Then
					pathAttachments = Mid(pathAttachments, 1, pathAttachments.Length - 1)
				End If
				Dim InizioRicerca As String = ""
				Dim DataRicercaMinima As Date = "26/02/1972 13:45:00"

				Sql = "SELECT * From [Generale].[dbo].[Utenti] A " &
					"Left Join [Generale].[dbo].[UtentiMails] B On A.idUtente = B.idUtente " &
					"Where A.idAnno = " & idAnno & " And A.idUtente = " & idUtente
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ok = False
						Ritorno = StringaErrore & " Nessun utente rilevato"
					Else
						Utenza = Rec("Mail").Value
						Password = Rec("PwdMail").Value
						UltimoControllo = "" & Rec("UltimoControllo").Value

						If UltimoControllo <> "" Then
							Dim mesi() As String = {"", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
							Dim DataSplit() As String = UltimoControllo.Split(";")
							Dim OrarioRicerca As String = DataSplit(0) & "-" & DataSplit(1) & "-" & DataSplit(2) & " " & DataSplit(3) & ":" & DataSplit(4) & ":" & DataSplit(5)
							DataRicercaMinima = OrarioRicerca

							InizioRicerca = "SENTSINCE " & DataSplit(2) & "-" & mesi(Val(DataSplit(1))) & "-" & DataSplit(0)
						End If
					End If
					Rec.Close()
				End If

				If Ok Then
					If TipoDB = "SQLSERVER" Then
						Sql = "SELECT IsNull(Max(idMail),0)+1 From Mails"
					Else
						Sql = "SELECT Coalesce(Max(idMail),0)+1 From Mails"
					End If
					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					'If Rec(0).Value Is DBNull.Value Then
					'	idMail = 1
					'Else
					idMail = Rec(0).Value
					'End If
					Rec.Close()

					Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
					Ritorno = Conn.EsegueSql(MP, Sql, Connessione)

					Dim imap As ImapClient

					imap = New ImapClient("imaps.aruba.it", 993, Utenza, Password, True)
					imap.Connect()
					imap.Authenticate()

					Try
						Dim recentMessages As List(Of MailMessage)
						'Dim mailBoxes As Mailbox() = imap.ListMailboxes(False)
						imap.SelectMailbox(Casella)
						If InizioRicerca = "" Then
							Dim quantiMessaggi As Integer = imap.GetMessageCount()
							recentMessages = imap.GetMessages(Casella, quantiMessaggi)
						Else
							recentMessages = imap.Search(InizioRicerca)
						End If
						' Dim recentMessages As List(Of MailMessage) = imap.Search("TEXT ""a""")
						For Each m As MailMessage In recentMessages
							Dim Prosegui As Boolean = False

							Dim dataMessaggio As Date = m.Date
							If DataRicercaMinima.Year <> 1972 Then
								If dataMessaggio > DataRicercaMinima Then
									Prosegui = True
								End If
							Else
								Prosegui = True
							End If

							If Prosegui Then
								Dim idMessage As String = SistemaStringa(m.MessageId)
								If idMessage.Length > 100 Then idMessage = Mid(idMessage, 1, 97) & "..."
								If idMessage = "" Then
									idMessage = "idAutomatico_" & m.Date.ToString.Replace(" ", "_")
								End If

								Sql = "Select * From Mails Where id = '" & idMessage & "'"
								Rec = Conn.LeggeQuery(MP, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
									Ok = False
								Else
									If Rec.Eof() Then
										Dim Destinatari As MailAddressCollection = m.To
										Dim Mittente As MailAddress = m.From

										'Ritorno &= m.MessageId & ";"

										'Ritorno &= SistemaStringa(Mittente.DisplayName) & ";"
										'Ritorno &= ";"
										'Ritorno &= SistemaStringa(Mittente.Address) & ";"
										' From
										' To
										'Ritorno &= SistemaStringa(m.Subject) & ";"
										'Ritorno &= SistemaStringa(m.Body) & ";"
										'Ritorno &= SistemaStringa(m.Date) & ";"
										'Ritorno &= ";"
										'Ritorno &= ";"
										'Ritorno &= ";"
										'' Attachments
										'Ritorno &= ";"
										'Ritorno &= ";"
										Sql = "Insert Into Mails Values (" &
											" " & idMail & ", " &
											" " & idUtente & ", " &
											"'" & SistemaStringa(m.Subject) & "', " &
											"'" & SistemaStringa(m.Body) & "', " &
											"'" & SistemaStringa(m.Date) & "', " &
											"'N', " &
											"'N', " &
											"'N', " &
											"'" & IIf(m.Attachments.Count > 0, "S", "N") & "', " &
											"'" & Casella & "', " &
											"'N', " &
											"'" & idMessage & "', " &
											"'" & m.From.Address & "', " &
											"'" & m.From.DisplayName & "' " &
											")"
										Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If

										If Ok Then
											Dim Attachments As AttachmentCollection = m.Attachments
											Dim Progressivo As Integer = 1

											'Ritorno &= "|"
											For Each a As Attachment In Attachments
												'Ritorno &= a.MediaType & ";"
												'Ritorno &= a.Name & ";"
												'Ritorno &= ";"
												'Ritorno &= ";"
												'Ritorno &= ";"

												Dim nomeFile As String = pathAttachments & "\" & idUtente & "\" & idMail & "-" & a.Name
												gf.CreaDirectoryDaPercorso(nomeFile)

												Using fileStream = New FileStream(nomeFile, FileMode.Create, FileAccess.Write)
													a.ContentStream.CopyTo(fileStream)
												End Using

												' Dim Buffer As new Stream = a.ContentStream

												Sql = "Insert Into MailsAttachment Values (" &
													" " & idMail & ", " &
													" " & Progressivo & ", " &
													"'" & SistemaStringa(a.MediaType) & "', " &
													"'" & SistemaStringa(a.Name) & "', " &
													"'" & "', " &
													"'', " &
													" " & a.ContentStream.Length & " " &
													")"
												Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
													Exit For
												End If

												Progressivo += 1
											Next

											'Ritorno &= "|"
										End If

										If Ok Then
											Dim Progressivo As Integer = 1

											For Each d As MailAddress In Destinatari
												'Ritorno &= SistemaStringa(d.DisplayName) & ";"
												'Ritorno &= ";"
												'Ritorno &= SistemaStringa(d.Address) & ";"

												Dim nome As String = "" & d.DisplayName
												If nome.Trim() = "" Then nome = d.Address


												Sql = "Insert Into MailsTo Values (" &
													" " & idMail & ", " &
													" " & Progressivo & ", " &
													"-1, " &
													"'" & SistemaStringa(nome) & "', " &
													"'" & SistemaStringa(d.Address) & "' " &
													")"
												Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
													Exit For
												End If
												Progressivo += 1
											Next
										End If

										'Ritorno &= "§"

										idMail += 1
									End If
								End If
								If Not Ok Then
									Exit For
								Else
									QuanteMails += 1
								End If
							End If
						Next
					Catch ex As Exception
						Ok = False
						Ritorno = StringaErrore & " " & ex.Message
					End Try

					imap.Disconnect()

					If Ok Then
						Dim Datella As String = Now.Year & ";" & Format(Now.Month, "00") & ";" & Format(Now.Day, "00") & ";" & Format(Now.Hour, "00") & ";" & Format(Now.Minute, "00") & ";" & Format(Now.Second, "00")

						Sql = "Update [Generale].[dbo].[UtentiMails] Set UltimoControllo = '" & Datella & "' Where idUtente = " & idUtente
						Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						Sql = "commit"
						Dim Ritorno2 As String = Conn.EsegueSql(MP, Sql, Connessione)
						Ritorno = QuanteMails
					Else
						Sql = "rollback"
						Dim Ritorno2 As String = Conn.EsegueSql(MP, Sql, Connessione)
					End If
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function
End Class
