﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Diagnostics.Eventing.Reader

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsMail
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaMails(Squadra As String, idAnno As String, idUtente As String, Folder As String, Filter As String, Label As String, SoloNuove As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				'If SoloNuove = "S" Then
				'	Dim m As New mailImap
				'	Dim Ritorno2 As String = m.RitornaMessaggi(Squadra, idAnno, idUtente, Folder)
				'End If

				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Altro As String = ""
				Dim Cosa As String = "*"

				'If Folder <> "" Then
				'	Altro &= " And Folder='" & Folder.Replace("'", "''") & "'"
				'End If
				If Filter <> "" Then
					If Filter = "Preferiti" Then
						Altro &= " And starred = 'S'"
					Else
						Altro &= " And important = 'S'"
					End If
				End If
				If Label <> "" Then
					Altro &= " And Label Like '%" & Label.Replace("'", "''") & "%'"
				End If
				If SoloNuove = "S" Then
					Altro = " And letto = 'N' And folder = 'Inbox'"
					Cosa = "Count(*)"
				Else
					Altro &= " Order By cast(substring(time,7,4) + substring(time,4,2) + substring(time,1,2) + substring(time,12,2) + substring(time,15,2) + substring(time,18,2) as numeric(15)) Desc"
				End If

				Sql = "SELECT " & Cosa & " From Mails " &
					"Where Eliminata = 'N' And idUtente=" & idUtente & Altro
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna mail ritornata"
					Else
						If SoloNuove = "" Or SoloNuove = "N" Then
							Do Until Rec.Eof
								Dim idMail As Integer = Rec("idMail").Value
								Dim Destinatari As String = ""
								Dim AttachMents As String = ""
								Dim Labels As String = ""

								Sql = "Select * From MailsTo Where idMail=" & idMail & " Order By progressivo"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								Do Until Rec2.Eof
									Destinatari &= Rec2("idUtente").Value & "^" & Rec2("name").value & "^" & Rec2("email").value & "°"

									Rec2.MoveNext
								Loop
								Rec2.Close

								Sql = "Select * From MailsAttachment Where idMail=" & idMail & " Order By progressivo"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								Do Until Rec2.Eof
									AttachMents &= Rec2("type").Value & "^" & Rec2("filename").Value & "^" & Rec2("preview").Value & "^" & Rec2("url").Value & "^" & Rec2("size").Value & "°"

									Rec2.MoveNext
								Loop
								Rec2.Close

								Sql = "Select * From MailsLabels Where idMail=" & idMail & " Order By progressivo"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								Do Until Rec2.Eof
									Labels &= Rec2("label").Value & "°"

									Rec2.MoveNext
								Loop
								Rec2.Close

								Ritorno &= idMail & ";"
								Ritorno &= Rec("subject").Value.replace(";", "***PV***") & ";"
								Ritorno &= Rec("message").Value.replace(";", "***PV***") & ";"
								Ritorno &= Rec("time").Value & ";"
								Ritorno &= Rec("letto").Value & ";"
								Ritorno &= Rec("starred").Value & ";"
								Ritorno &= Rec("important").Value & ";"
								Ritorno &= Rec("hasAttachments").Value & ";"
								Ritorno &= Rec("folder").Value & ";"
								Ritorno &= Destinatari & ";"
								Ritorno &= AttachMents & ";"
								Ritorno &= Labels & ";"
								Ritorno &= Rec("Mittente").Value & ";"
								Ritorno &= Rec("NomeMittente").Value & ";"
								Ritorno &= "§"

								Rec.MoveNext
							Loop
						Else
							If Rec(0).Value Is DBNull.Value Then
								Ritorno = 0
							Else
								Ritorno = Rec(0).Value
							End If
						End If
					End If
					Rec.Close()
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiungeMail(Squadra As String, idUtente As String, from As String, subject As String, message As String, time As String,
								 letto As String, starred As String, important As String, hasAttachments As String, folder As String,
								 mailsTo As String, attachments As String, labels As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim ma As New mail
			Dim Mittente As String = ""
			Dim mailMittente As String = ""
			Dim gf As New GestioneFilesDirectory
			Dim righe As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
			Dim c() As String = righe.Split(";")
			Dim pathFisico As String = c(0).Trim.Replace(vbCrLf, "")
			If pathFisico.EndsWith("\") Then
				pathFisico = Mid(pathFisico, 1, pathFisico.Length - 1)
			End If
			Dim urlFisico As String = c(2).Trim.Replace(vbCrLf, "")
			If urlFisico.EndsWith("/") Then
				urlFisico = Mid(urlFisico, 1, urlFisico.Length - 1)
			End If
			urlFisico = urlFisico.Replace("Multimedia", "Allegati")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim idMail As Integer = -1
				Dim Ok As Boolean = True
				Dim Destinatari As String = ""

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Sql = "SELECT Max(idMail)+1 From Mails"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If Rec(0).Value Is DBNull.Value Then
					idMail = 1
				Else
					idMail = Rec(0).Value
				End If
				Rec.Close

				Sql = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If Rec.Eof Then
					Ritorno = StringaErrore & " Nessun utente rilevato"
					Ok = False
				Else
					Mittente = Rec("Cognome").Value & " " & Rec("Nome").Value
					mailMittente = Rec("EMail").Value
				End If
				Rec.Close

				If Ok Then
					Dim sTo() As String = mailsTo.Split("§")

					' Scrive i dati della mail propria mail nella casella Inviate
					Sql = "Insert Into Mails Values (" &
						" " & idMail & ", " &
						" " & idUtente & ", " &
						"'" & subject.Replace("'", "''") & "', " &
						"'" & message.Replace("'", "''") & "', " &
						"'" & time & "', " &
						"'" & letto & "', " &
						"'" & starred & "', " &
						"'" & important & "', " &
						"'" & hasAttachments & "', " &
						"'Inviate', " &
						"'N', " &
						"-1, " &
						"'" & Mittente.Replace("'", "''") & "'," &
						"'" & mailMittente.Replace("'", "''") & "'" &
						")"
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						'Ritorno = ma.SendEmail(Squadra, from, subject, message, from, {""})
						'If Ritorno <> "*" Then
						'	Ok = False
						'End If

						Dim Progressivo As Integer = 0

						For Each t2 As String In sTo
							If t2 <> "" Then
								Dim c2() As String = t2.Split(";")

								Progressivo += 1
								Sql = "Insert Into MailsTo Values (" &
									" " & idMail & ", " &
									" " & Progressivo & ", " &
									" " & c2(0) & ", " &
									"'" & c2(1).Replace("'", "''") & "', " &
									"'" & c2(2).Replace("'", "''") & "' " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
									Exit For
								Else
									Destinatari &= c2(2) & ";"
								End If
							End If
						Next
					End If

					If Ok Then
						Dim sAt() As String = attachments.Split("§")
						Dim Progressivo As Integer = 0

						For Each t2 As String In sAt
							If t2 <> "" Then
								Dim c2() As String = t2.Split(";")
								Dim Type As String = gf.TornaEstensioneFileDaPath(c2(0))
								Dim Size As String = c2(1)

								Progressivo += 1
								Sql = "Insert Into MailsAttachment Values (" &
									" " & idMail & ", " &
									" " & Progressivo & ", " &
									"'" & Type & "', " &
									"'" & (pathFisico & "\" & c2(0)).Replace("'", "''") & "', " &
									"'" & "', " &
									"'" & (urlFisico & "/" & c2(0).Replace("\", "/")).Replace("'", "''") & "', " &
									" " & Size & " " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
									Exit For
								End If
							End If
						Next
					End If

					If Ok Then
						Dim sLab() As String = labels.Split("§")
						Dim Progressivo As Integer = 0

						For Each t2 As String In sLab
							If t2 <> "" Then
								Progressivo += 1
								Sql = "Insert Into MailsLabels Values (" &
									" " & idMail & ", " &
									" " & Progressivo & ", " &
									"'" & t2.Replace("'", "''") & "' " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
									Exit For
								End If
							End If
						Next
					End If
				End If

				If Ok Then
					Dim Dests() As String = Destinatari.Split(";")
					Dim attach() As String = attachments.Split("§")
					Dim aa As String = ""

					For Each a As String In attach
						If a <> "" Then
							Dim aaa() As String = a.Split(";")
							aa &= pathFisico & "\" & a(0) & ";"
						End If
					Next

					Dim attachs() As String = aa.Split(";")
					For Each d As String In Dests
						Ritorno = ma.SendEmail(Squadra, from, subject, message, d, attachs)
						If Ritorno <> "*" Then
							Ok = False
							Exit For
						End If
					Next
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
	Public Function RitornaDestinatari(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

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
				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = Str(Val(c(1))).Trim

				Sql = "Select * From Utenti Where idSquadra=" & c(1) & " And Eliminato='N' Order By Cognome, Nome"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("idUtente").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("EMail").Value & ";" & Rec("idTipologia").Value & "§"

						Rec.MoveNext
					Loop
					Rec.Close
				End If

				Ritorno &= "|"
				Sql = "Select * From [" & Squadra & "].[dbo].[Categorie] Where Eliminato='N' Order By Descrizione"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Dim Partecipanti As String = ""

						Sql = "Select * From [" & Squadra & "].[dbo].[UtentiCategorie] Where idCategoria=" & Rec("idCategoria").Value
						Rec2 = LeggeQuery(Conn, Sql, Connessione)
						Do Until Rec2.Eof
							Partecipanti &= Rec2("idUtente").Value & "^"

							Rec2.MoveNext
						Loop
						Rec2.Close
						If Partecipanti <> "" Then
							Partecipanti = Mid(Partecipanti, 1, Partecipanti.Length - 1)
						End If

						Ritorno &= Rec("idCategoria").Value & ";" & Rec("Descrizione").Value & ";" & Partecipanti & "§"

						Rec.MoveNext
					Loop
					Rec.Close
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaMsg(Squadra As String, idUtente As String, idMail As String) As String
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

				Sql = "Select * From Mails Where idMail=" & idMail & " And idUtente=" & idUtente
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna mail rilevata"
					Else
						Dim statoAttuale As String = "" & Rec("folder").Value
						Rec.Close

						If statoAttuale = "Eliminate" Then
							Sql = "Update Mails Set Eliminata='S' Where idMail=" & idMail & " And idUtente=" & idUtente
						Else
							Sql = "Update Mails Set folder = 'Eliminate' Where idMail=" & idMail & " And idUtente=" & idUtente
						End If

						Ritorno = EsegueSql(Conn, Sql, Connessione)
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ImpostaMsgComeLetto(Squadra As String, idUtente As String, idMail As String) As String
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

				Sql = "Update Mails Set Letto='S' Where idMail=" & idMail & " And idUtente=" & idUtente
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ImpostaFolderMsg(Squadra As String, idUtente As String, idMail As String, Folder As String) As String
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

				Sql = "Update Mails Set folder='" & Folder & "' Where idMail=" & idMail & " And idUtente=" & idUtente
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ImpostaMsgComePreferito(Squadra As String, idUtente As String, idMail As String, Preferito As String) As String
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

				Sql = "Update Mails Set starred='" & Preferito & "' Where idMail=" & idMail & " And idUtente=" & idUtente
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ImpostaMsgComeImportante(Squadra As String, idUtente As String, idMail As String, Importante As String) As String
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

				Sql = "Update Mails Set important='" & Importante & "' Where idMail=" & idMail & " And idUtente=" & idUtente
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function
End Class