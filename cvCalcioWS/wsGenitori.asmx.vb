Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://genitori.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsGenitori
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaIdGiocatore(Squadra As String, idUtente As String) As String
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
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As Integer = -1
				Dim Ok As Boolean = True

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value

						Dim gf As New GestioneFilesDirectory
						Dim path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
						Dim pp() As String = path.Split(";")
						If Strings.Right(pp(0), 1) <> "\" Then
							pp(0) = pp(0) & "\"
						End If
						Dim a() As String = Squadra.Split("_")
						Dim Anno As Integer = Val(a(0))
						Dim p As String = pp(0) & Squadra & "\Certificati\Anno" & Anno & "\" & idGiocatore & "\"
						gf.ScansionaDirectorySingola(p)
						Dim filetti() As String = gf.RitornaFilesRilevati
						Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

						Ritorno = idGiocatore & ";"
						For i As Integer = 1 To qFiletti
							Ritorno &= gf.TornaNomeFileDaPath(filetti(i)) & ";"
						Next
					End If
					Rec.Close
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMails(Squadra As String, idUtente As String) As String
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
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As Integer = -1
				Dim Ok As Boolean = True

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value
					End If
					Rec.Close
				End If

				If Ok Then
					Sql = "Select A.*, B.Cognome + ' ' + B.Nome As Giocatore From GiocatoriDettaglio A " &
						"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
						"Where A.idGiocatore=" & idGiocatore
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
							Ok = False
						Else
							Dim Genitore1 As String = Rec("Genitore1").Value
							Dim Genitore2 As String = Rec("Genitore2").Value
							Dim Giocatore As String = Rec("Giocatore").Value
							Dim mail1 As String = Rec("MailGenitore1").Value
							Dim mail2 As String = Rec("MailGenitore2").Value
							Dim mail3 As String = Rec("MailGenitore3").Value
							Rec.Close

							Dim Attiva1 As String = "N"
							Dim Attiva2 As String = "N"
							Dim Attiva3 As String = "N"

							Sql = "Select * From GiocatoriMails Where idGiocatore=" & idGiocatore
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
							Else
								If Not Rec.Eof() Then
									Do Until Rec.Eof
										Select Case Val(Rec("Progressivo").Value)
											Case 1
												Attiva1 = Rec("Attiva").Value
											Case 2
												Attiva2 = Rec("Attiva").Value
											Case 3
												Attiva3 = Rec("Attiva").Value
										End Select

										Rec.MoveNext
									Loop
									Rec.Close

									Ritorno = Genitore1 & ";" & mail1 & ";" & Attiva1 & ";§"
									Ritorno &= Genitore2 & ";" & mail2 & ";" & Attiva2 & ";§"
									Ritorno &= Giocatore & ";" & mail3 & ";" & Attiva3 & ";§"
								End If
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaMails(Squadra As String, idUtente As String, Attiva1 As String, Attiva2 As String, Attiva3 As String) As String
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
				Dim Sql As String = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
				Dim idGiocatore As Integer = -1
				Dim Ok As Boolean = True

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
						Ok = False
					Else
						idGiocatore = Rec("idGiocatore").Value
					End If
					Rec.Close
				End If

				If Ok Then
					If Not Ritorno.Contains(StringaErrore) Then
						Sql = "Update GiocatoriMails Set Attiva = '" & Attiva1 & "' Where idGiocatore=" & idGiocatore & " And Progressivo = 1"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

						Sql = "Update GiocatoriMails Set Attiva = '" & Attiva2 & "' Where idGiocatore=" & idGiocatore & " And Progressivo = 2"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

						Sql = "Update GiocatoriMails Set Attiva = '" & Attiva3 & "' Where idGiocatore=" & idGiocatore & " And Progressivo = 3"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

						Ritorno = "*"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function
End Class