Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Drawing
Imports System.Threading.Tasks

<System.Web.Services.WebService(Namespace:="http://cvcalcio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsGenerale
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function LeggeMailbox() As String
		Dim m As New mailImap
		Dim ritorno As string = m.RitornaMessaggi("0001_00002", "1", "1", "Inbox")

		Return ritorno
	End Function

	<WebMethod()>
	Public Function AggiungeCFaCSV(NomeFile As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Tutto As String = gf.LeggeFileIntero(NomeFile)
		Dim Righe() As String = Tutto.Split(vbCrLf)
		Dim cf As New CodiceFiscale
		Dim Righelle As Integer = 0

		gf.ApreFileDiTestoPerScrittura("C:\Users\looigi\Desktop\NuovoFileIscritti.csv")
		For Each r As String In Righe
			If r <> "" Then
				r = r.Replace(vbLf, "")
				Dim campi() As String = r.Split(";")
				If campi.Length > 4 Then
					Dim DataNascita As String = campi(3).Trim & "-" & campi(4).Trim.Replace("-", "")
					DataNascita = DataNascita.ToUpper.Replace("GEN", "01")
					DataNascita = DataNascita.ToUpper.Replace("FEB", "02")
					DataNascita = DataNascita.ToUpper.Replace("MAR", "03")
					DataNascita = DataNascita.ToUpper.Replace("APR", "04")
					DataNascita = DataNascita.ToUpper.Replace("MAG", "05")
					DataNascita = DataNascita.ToUpper.Replace("GIU", "06")
					DataNascita = DataNascita.ToUpper.Replace("LUG", "07")
					DataNascita = DataNascita.ToUpper.Replace("AGO", "08")
					DataNascita = DataNascita.ToUpper.Replace("SET", "09")
					DataNascita = DataNascita.ToUpper.Replace("OTT", "10")
					DataNascita = DataNascita.ToUpper.Replace("NOV", "11")
					DataNascita = DataNascita.ToUpper.Replace("DIC", "12")
					DataNascita = DataNascita.ToUpper.Replace("/", "-")
					DataNascita = DataNascita.ToUpper.Replace("--", "-")
					Dim Nome As String = campi(1)
					Dim Cognome As String = campi(2)
					Dim Comune As String = campi(5)
					Dim scf As String = ""
					If DataNascita.Length > 6 Then
						scf = cf.CreaCodiceFiscale(Cognome, Nome, DataNascita, Comune, True)
					Else
						DataNascita = ""
						scf = ""
					End If
					If scf.Length > 10 Or scf = "" Then
						Dim nuovaRiga As String = Cognome & ";" & Nome & ";" & Comune & ";" & DataNascita & ";" & scf & ";"
						gf.ScriveTestoSuFileAperto(nuovaRiga)
						Righelle += 1
					End If
				End If
			End If
		Next
		gf.ChiudeFileDiTestoDopoScrittura()

		Return Righelle
	End Function

	<WebMethod()>
	Public Function TestDataMaggiorenne(Valore As String) As String
		Dim d() As String = Valore.Split("/")
		Valore = d(2) & "-" & d(1) & "-" & d(0)
		Dim dd As Date = Convert.ToDateTime(Valore)
		Dim Oggi As Date = Now
		Dim diff As Integer = DateDiff(DateInterval.Year, dd, Oggi)

		Return diff
	End Function

	<WebMethod()>
	Public Function ProvaConversioneValore(valore As String) As String
		Dim m As New mail
		Dim Ritorno As String = convertNumberToReadableString(valore)
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMeteoWS() As String
		Dim m As New mail
		Dim Ritorno As String = RitornaMeteo("41.89", "12.48")
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InviaMail(Oggetto As String, Body As String, ChiRiceve As String) As String
		Dim m As New mail
		Dim Ritorno As String = m.SendEmail("", "looigi@gmail.com", Oggetto, Body, ChiRiceve, {"E:\Sorgenti\VB.Net\Miei\WEB\Webservices\cvCalcio\cvCalcioWS\cvCalcioWS\Impostazioni\Paths.txt", "E:\Sorgenti\VB.Net\Miei\WEB\Webservices\cvCalcio\cvCalcioWS\cvCalcioWS\Impostazioni\PercorsoSito.txt"})
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InviaMailConAllegato(Squadra As String, Oggetto As String, Body As String, Destinatario As String, Allegato As String, AllegatoOMultimedia As String, Mittente As String) As String
		Dim m As New mail
		Dim Ritorno As String = m.SendEmail(Squadra, Mittente, Oggetto, Body, Destinatario, {Allegato.Replace("/", "\")}, AllegatoOMultimedia)
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InviaSollecitoPagamento(Squadra As String, Destinatario As String, Dati As String, Mittente As String) As String
		Dim m As New mail

		Dim Oggetto As String = "Sollecito pagamento"
		Dim d() As String = Dati.Split(";")
		Dim gf As New GestioneFilesDirectory
		Dim Body As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\mail_sollecito.txt")
		Body = Body.Replace("***Data scadenza tab rate***", d(0))
		Body = Body.Replace("***Descrizione tab rate ***", d(1))
		Body = Body.Replace("***Importo tab rate***", d(2))
		Body = Body.Replace("****cognome menu&nbsp; anagrafica3***", d(3))
		Body = Body.Replace("***Nome menu anagrafica3", d(4))
		Body = Body.Replace("***nome societ&agrave; menu settaggi***", d(5))

		' Dim Body As String = "In data " & d(0) & " è scaduta la rata '" & d(1) & "' dell'importo di Euro " & d(2) & ".<br />Si prega di passare urgentemente in segreteria.<br />Grazie"

		Dim Ritorno As String = m.SendEmail(Squadra, Mittente, Oggetto, Body, Destinatario, {""}, "")
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SistemaImmagini(Squadra As String, idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim gf As New GestioneFilesDirectory
				Dim PathBase As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
				PathBase = PathBase.Replace(vbCrLf, "")
				PathBase = PathBase.Replace(Chr(13), "")
				PathBase = PathBase.Replace(Chr(10), "")
				If Strings.Right(PathBase, 1) = "\" Then
					PathBase = Mid(PathBase, 1, PathBase.Length - 1)
				End If
				PathBase &= "\" & Squadra.Replace(" ", "_")

				Dim Sql As String = "Select idGiocatore From Giocatori Where idAnno=" & idAnno
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sconosciuto As String = PathBase & "\Giocatori\Sconosciuto.png"
				Dim Ok As Boolean = True
				Dim Aggiunte As Integer = 0
				Dim Eliminate As Integer = 0

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					Do Until Rec.Eof
						Dim Percorso As String = PathBase & "\Giocatori\" & idAnno & "_" & Rec("idGiocatore").Value.ToString & ".jpg"
						If Not File.Exists(Percorso) Then
							Try
								File.Copy(Sconosciuto, Percorso)
								Aggiunte += 1
							Catch ex As Exception
								Ritorno = StringaErrore & ex.Message
								Ok = False
								Exit Do
							End Try
						End If

						Rec.MoveNext()
					Loop
					Rec.Close()
				End If

				If Ok Then
					Sql = "Select idAllenatore From Allenatori Where idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof
							Dim Percorso As String = PathBase & "\Allenatori\" & idAnno & "_" & Rec("idAllenatore").Value.ToString & ".jpg"
							If Not File.Exists(Percorso) Then
								Try
									File.Copy(Sconosciuto, Percorso)
									Aggiunte += 1
								Catch ex As Exception
									Ritorno = StringaErrore & ex.Message
									Ok = False
									Exit Do
								End Try
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				If Ok Then
					Sql = "Select idArbitro From Arbitri Where idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof
							Dim Percorso As String = PathBase & "\Arbitri\" & Rec("idArbitro").Value.ToString & ".jpg"
							If Not File.Exists(Percorso) Then
								Try
									File.Copy(Sconosciuto, Percorso)
									Aggiunte += 1
								Catch ex As Exception
									Ritorno = StringaErrore & ex.Message
									Ok = False
									Exit Do
								End Try
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				If Ok Then
					Sql = "Select idAvversario From SquadreAvversarie"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof
							Dim Percorso As String = PathBase & "\Avversari\" & Rec("idAvversario").Value.ToString & ".jpg"
							If Not File.Exists(Percorso) Then
								Try
									File.Copy(Sconosciuto, Percorso)
									Aggiunte += 1
								Catch ex As Exception
									Ritorno = StringaErrore & ex.Message
									Ok = False
									Exit Do
								End Try
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				If Ok Then
					Sql = "Select idCategoria From Categorie Where idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof
							Dim Percorso As String = PathBase & "\Categorie\" & idAnno & "_" & Rec("idCategoria").Value.ToString & ".jpg"
							If Not File.Exists(Percorso) Then
								Try
									File.Copy(Sconosciuto, Percorso)
									Aggiunte += 1
								Catch ex As Exception
									Ritorno = StringaErrore & ex.Message
									Ok = False
									Exit Do
								End Try
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				If Ok Then
					Sql = "Select idDirigente From Dirigenti Where idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof
							Dim Percorso As String = PathBase & "\Dirigenti\" & idAnno & "_" & Rec("idDirigente").Value.ToString & ".jpg"
							If Not File.Exists(Percorso) Then
								Try
									File.Copy(Sconosciuto, Percorso)
									Aggiunte += 1
								Catch ex As Exception
									Ritorno = StringaErrore & ex.Message
									Ok = False
									Exit Do
								End Try
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				If Ok Then
					Sql = "Select idUtente From Utenti Where idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof
							Dim Percorso As String = PathBase & "\Utenti\" & idAnno & "_" & Rec("idUtente").Value.ToString & ".jpg"
							If Not File.Exists(Percorso) Then
								Try
									File.Copy(Sconosciuto, Percorso)
									Aggiunte += 1
								Catch ex As Exception
									Ritorno = StringaErrore & ex.Message
									Ok = False
									Exit Do
								End Try
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				If Ok Then
					Sql = "Select idUtente From UtentiMobile Where idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Ok = False
					Else
						Do Until Rec.Eof
							Dim Percorso As String = PathBase & "\UtentiMobile\" & idAnno & "_" & Rec("idUtente").Value.ToString & ".jpg"
							If Not File.Exists(Percorso) Then
								Try
									File.Copy(Sconosciuto, Percorso)
									Aggiunte += 1
								Catch ex As Exception
									Ritorno = StringaErrore & ex.Message
									Ok = False
									Exit Do
								End Try
							End If

							Rec.MoveNext()
						Loop
						Rec.Close()
					End If
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\Allenatori"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						Dim c() As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "").Split("_")
						Sql = "Select * From Allenatori Where idAnno=" & c(0) & " And idAllenatore=" & c(1)
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
							Exit For
						Else
							If Rec.Eof Then
								Try
									Kill(Filetti(i))
									Eliminate += 1
								Catch ex As Exception
									Ok = False
									Ritorno = StringaErrore & ex.Message
								End Try
							End If
							Rec.Close
						End If
					Next
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\Arbitri"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						Dim c As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "")
						Sql = "Select * From Arbitri Where idAnno=" & idAnno & " And idArbitro=" & c
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
							Exit For
						Else
							If Rec.Eof Then
								Try
									Kill(Filetti(i))
									Eliminate += 1
								Catch ex As Exception
									Ok = False
									Ritorno = StringaErrore & ex.Message
								End Try
							End If
							Rec.Close
						End If
					Next
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\Avversari"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						Dim c As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "")
						Sql = "Select * From SquadreAvversarie Where idAvversario=" & c
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
							Exit For
						Else
							If Rec.Eof Then
								Try
									Kill(Filetti(i))
									Eliminate += 1
								Catch ex As Exception
									Ok = False
									Ritorno = StringaErrore & ex.Message
								End Try
							End If
							Rec.Close
						End If
					Next
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\Categorie"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						Dim c() As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "").Split("_")
						Sql = "Select * From Categorie Where idAnno=" & c(0) & " And idCategoria=" & c(1)
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
							Exit For
						Else
							If Rec.Eof Then
								Try
									Kill(Filetti(i))
									Eliminate += 1
								Catch ex As Exception
									Ok = False
									Ritorno = StringaErrore & ex.Message
								End Try
							End If
							Rec.Close
						End If
					Next
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\Dirigenti"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						Dim c() As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "").Split("_")
						Sql = "Select * From Dirigenti Where idAnno=" & c(0) & " And idDirigente=" & c(1)
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
							Exit For
						Else
							If Rec.Eof Then
								Try
									Kill(Filetti(i))
									Eliminate += 1
								Catch ex As Exception
									Ok = False
									Ritorno = StringaErrore & ex.Message
								End Try
							End If
							Rec.Close
						End If
					Next
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\Giocatori"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						If Not Filetti(i).Contains("Sconosciuto") Then
							Dim c() As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "").Split("_")
							Sql = "Select * From Giocatori Where idAnno=" & c(0) & " And idGiocatore=" & c(1)
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
								Exit For
							Else
								If Rec.Eof Then
									Try
										Kill(Filetti(i))
										Eliminate += 1
									Catch ex As Exception
										Ok = False
										Ritorno = StringaErrore & ex.Message
									End Try
								End If
								Rec.Close
							End If
						End If
					Next
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\Utenti"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						Dim c() As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "").Split("_")
						Sql = "Select * From Utenti Where idAnno=" & c(0) & " And idUtente=" & c(1)
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
							Exit For
						Else
							If Rec.Eof Then
								Try
									Kill(Filetti(i))
									Eliminate += 1
								Catch ex As Exception
									Ok = False
									Ritorno = StringaErrore & ex.Message
								End Try
							End If
							Rec.Close
						End If
					Next
				End If

				If Ok Then
					Dim Percorso As String = PathBase & "\UtentiMobile"
					gf.ScansionaDirectorySingola(Percorso)
					Dim Filetti() As String = gf.RitornaFilesRilevati
					Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

					For i As Integer = 1 To qFiletti
						Dim c() As String = Filetti(i).Replace(Percorso & "\", "").ToUpper().Replace(".JPG", "").Split("_")
						Sql = "Select * From UtentiMobile Where idAnno=" & c(0) & " And idUtente=" & c(1)
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							Ok = False
							Exit For
						Else
							If Rec.Eof Then
								Try
									Kill(Filetti(i))
									Eliminate += 1
								Catch ex As Exception
									Ok = False
									Ritorno = StringaErrore & ex.Message
								End Try
							End If
							Rec.Close
						End If
					Next
				End If

				If Ok Then
					Ritorno = "Immagini Aggiunte: " & Aggiunte & " - Eliminate: " & Eliminate
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiornaDB(Squadra As String, ByVal Numero As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Try
					Select Case Numero
						Case "1"
							Sql = "Create Table CampiEsterni (idPartita Integer , Descrizione Text(255), CONSTRAINT TelefonatePK PRIMARY KEY (idPartita))"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Case "2"
							Sql = "Create Table CoordinatePartite (idPartita Integer, Lat Text(15), Lon Text(15), CONSTRAINT CoordPK PRIMARY KEY (idPartita))"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Case "3"
							Sql = "Create Table MeteoPartite (idPartita Integer, Tempo Text(30), Gradi Text(10), Umidita Text(10), Pressione Text(10), CONSTRAINT MeteoPK PRIMARY KEY (idPartita))"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Case "4"
							Sql = "Create Table AvversariCoord (idAvversario Integer, Lat Text(30), Lon Text(30), CONSTRAINT AvvCoordPK PRIMARY KEY (idAvversario))"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Case "5"
							Sql = "Alter Table CalendarioPartite Add Giocata Text(1)"
							Ritorno = EsegueSql(Conn, Sql, Connessione)

							Sql = "Update CalendarioPartite Set Giocata='S'"
							Ritorno = EsegueSql(Conn, Sql, Connessione)

							Sql = "Alter Table CalendarioDate Add idPartita Integer"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Case "6"
							Sql = "Create Table Giornata (idUtente Integer, idAnno Integer, idCategoria Integer, idGiornata Integer, CONSTRAINT GiornataPK PRIMARY KEY (idUtente, idAnno, idCategoria))"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Case "7"
							Sql = "Alter Table Anni Add NomeSquadra Text(50)"
							Ritorno = EsegueSql(Conn, Sql, Connessione)

							Sql = "Create Table AnnoAttualeUtenti (idUtente Integer, idAnno Integer, CONSTRAINT AnnoAttualeUtentiPK PRIMARY KEY (idUtente))"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						Case Else
							Ritorno = StringaErrore & " Valore non valido"
					End Select
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaVersioneApplicazione() As String
		Dim Ritorno As String = ""

		Dim gf As New GestioneFilesDirectory
		gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\NuoveVersioni\")
		Dim NuovaVersione As String = gf.LeggeFileIntero(Server.MapPath(".") & "\NuoveVersioni\Versione.txt")
		If NuovaVersione <> "" Then
			Ritorno = NuovaVersione
		Else
			Ritorno = StringaErrore & " Nessuna nuova versione rilevata"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaNumeroFattura(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From DatiFattura Where Anno = " & Now.Year

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = StringaErrore & " " & Rec
				Else
					If Rec.Eof Then
						Sql = "Insert Into DatiFattura Values (" & Now.Year & ", 0)"
						Ritorno = EsegueSql(Conn, Sql, Connessione)

						Ritorno = 0
					Else
						Ritorno = Rec("Progressivo").Value
					End If
					Rec.Close
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaNumeroFattura(Squadra As String, NumeroFattura As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = "Update DatiFattura Set Progressivo=" & NumeroFattura & " Where Anno=" & Now.Year
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaImpostazioni(Cod_Squadra As String, idAnno As String, Descrizione As String, NomeSquadra As String, Lat As String, Lon As String,
									  Indirizzo As String, CampoSquadra As String, NomePolisportiva As String, Mail As String, PEC As String,
									  Telefono As String, PIva As String, CodiceFiscale As String, CodiceUnivoco As String, SitoWeb As String, MittenteMail As String,
									  GestionePagamenti As String, CostoScuolaCalcio As String, idUtente As String, Widgets As String, Suffisso As String,
									  IscrFirmaEntrambi As String, ModuloAssociato As String, PercCashBack As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Cod_Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				If CostoScuolaCalcio = "" Then CostoScuolaCalcio = "0"

				Sql = "Update Anni Set " &
					"Descrizione = '" & Descrizione.Replace("'", "''") & "', " &
					"NomeSquadra = '" & NomeSquadra.Replace("_", " ").Replace("'", "''") & "', " &
					"Lat= '" & Lat & "', " &
					"Lon = '" & Lon & "', " &
					"Indirizzo = '" & Indirizzo.Replace("'", "''") & "', " &
					"CampoSquadra = '" & CampoSquadra.Replace("'", "''") & "', " &
					"NomePolisportiva = '" & NomePolisportiva.Replace("'", "''") & "', " &
					"Mail = '" & Mail.Replace("'", "''") & "', " &
					"PEC = '" & PEC.Replace("'", "''") & "', " &
					"Telefono = '" & Telefono.Replace("'", "''") & "', " &
					"PIva = '" & PIva.Replace("'", "''") & "', " &
					"CodiceFiscale = '" & CodiceFiscale.Replace("'", "''") & "', " &
					"CodiceUnivoco = '" & CodiceUnivoco.Replace("'", "''") & "', " &
					"SitoWeb = '" & SitoWeb.Replace("'", "''") & "', " &
					"MittenteMail = '" & MittenteMail.Replace("'", "''") & "', " &
					"GestionePagamenti = '" & GestionePagamenti & "', " &
					"CostoScuolaCalcio=" & CostoScuolaCalcio & ", " &
					"Suffisso='" & Suffisso & "', " &
					"iscrFirmaEntrambi='" & IscrFirmaEntrambi & "', " &
					"PercCashBack=" & PercCashBack & ", " &
					"ModuloAssociato='" & ModuloAssociato & "' " &
					"Where idAnno = " & idAnno
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Sql = "Update [Generale].[dbo].[Utenti] Set Widgets = '" & Widgets & "' Where idUtente=" & idUtente
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaImpostazioni(Squadra As String, idUtente As String) As String
		If Squadra = "" Then
			Return "*" ' StringaErrore & " Nessuna squadra impostata"
		End If

		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim ConnessioneGen As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim ConnGen As Object = ApreDB(ConnessioneGen)

			If TypeOf (Conn) Is String Or TypeOf (ConnGen) Is String Then
				If TypeOf (Conn) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
				Else
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnGen
				End If
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = Str(Val(c(1))).Trim
				Dim Anni As New List(Of Integer)
				Dim MeseAttivazione As New List(Of Integer)
				Dim AnnoAttivazione As New List(Of Integer)

				Sql = "Select * From SquadraAnni A " &
					"Left Join Squadre B On A.idSquadra = B.idSquadra " &
					"Where A.idSquadra=" & codSquadra & " Order By A.idAnno Desc"
				Rec = LeggeQuery(ConnGen, Sql, ConnessioneGen)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof()
						Anni.Add(Rec("idAnno").Value)
						MeseAttivazione.Add(Rec("MeseAttivazione").Value)
						AnnoAttivazione.Add(Rec("AnnoAttivazione").Value)

						Rec.MoveNext()
					Loop
					Rec.Close()

					If Anni.Count > 0 Then
						Sql = "Select * From Anni"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim NomeSquadra As String = Rec("NomeSquadra").Value
							Rec.Close()

							Ritorno = ""
							Dim quale As Integer = 0
							For Each a As Integer In Anni
								Dim sAnno As String = Format(a, "0000")
								Dim sCodSquadra As String = codSquadra.Trim
								While sCodSquadra.Length <> 5
									sCodSquadra = "0" & sCodSquadra
								End While
								Dim Codice As String = sAnno & "_" & sCodSquadra

								Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
								Dim P() As String = PathAllegati.Split(";")
								If Strings.Right(P(0), 1) = "\" Then
									P(0) = Mid(P(0), 1, P(0).Length - 1)
								End If
								Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
								pp = pp.Trim()
								If Strings.Right(pp, 1) = "\" Then
									pp = Mid(pp, 1, pp.Length - 1)
								End If
								Dim pathFirma1 As String = P(2) & "/" & NomeSquadra.Replace(" ", "_") & "/Utenti/" & Anno & "_" & idUtente & "_Firma.kgb"
								Dim urlFirma1 As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Utenti\" & Anno & "_" & idUtente & "_Firma.kgb"
								Dim esisteFirma As String = "N"

								If File.Exists(urlFirma1) Then
									esisteFirma = "S"
								End If

								Dim Widgets As String = ""

								Sql = "Select * From [Generale].[dbo].[Utenti] Where idUtente=" & idUtente
								Rec = LeggeQuery(ConnGen, Sql, ConnessioneGen)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									Widgets = "" & Rec("Widgets").Value
								End If
								Rec.Close()

								Sql = "Select A.*, B.idAvversario, B.idCampo " &
									"From [" & Codice & "].[dbo].[Anni] A Left Join [" & Codice & "].[dbo].[SquadreAvversarie] B On A.NomeSquadra = B.Descrizione " &
									"Order By idAnno Desc"
								Rec = LeggeQuery(ConnGen, Sql, ConnessioneGen)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									' Ritorno = ""
									Do Until Rec.Eof
										Ritorno &= Rec("idAnno").Value & ";" &
											Rec("Descrizione").Value & ";" &
											Rec("NomeSquadra").Value & ";" &
											Rec("Lat").Value & ";" &
											Rec("Lon").Value & ";" &
											Rec("Indirizzo").Value & ";" &
											Rec("CampoSquadra").Value & ";" &
											Rec("NomePolisportiva").Value & ";" &
											Rec("idAvversario").Value & ";" &
											Rec("idCampo").Value & ";" &
											Rec("Mail").Value & ";" &
											Rec("PEC").Value & ";" &
											Rec("Telefono").Value & ";" &
											Rec("PIva").Value & ";" &
											Rec("CodiceFiscale").Value & ";" &
											Rec("CodiceUnivoco").Value & ";" &
											Rec("SitoWeb").Value & ";" &
											Rec("MittenteMail").Value & ";" &
											Rec("GestionePagamenti").Value & ";" &
											esisteFirma & ";" &
											pathFirma1 & ";" &
											MeseAttivazione.Item(quale) & ";" &
											AnnoAttivazione.Item(quale) & ";" &
											Rec("CostoScuolaCalcio").value & ";" &
											Widgets & ";" &
											Rec("Suffisso").Value & ";" &
											Rec("iscrFirmaEntrambi").Value & ";" &
											Rec("ModuloAssociato").Value & ";" &
											Rec("PercCashBack").Value & ";" &
											"§"

										Rec.MoveNext()
									Loop
									Rec.Close()
								End If
								quale += 1
							Next
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaAnni(Squadra As String) As String
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
					Sql = "SELECT * FROM Anni"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Ritorno = ""
						Do Until Rec.Eof
							Ritorno &= Rec("idAnno").Value & ";" &
								Rec("Descrizione").Value & ";" &
								Rec("NomeSquadra").Value & ";" &
								"§"
							Rec.MoveNext()
						Loop
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
	Public Function RitornaValoriPerRegistrazione(Squadra As String) As String
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
				Dim Anno As String

				If Now.Month >= 8 Then
					' Ci si sta registrando per l'anno in corso
					Anno = Now.Year.ToString.Trim
				Else
					' Ci si sta registrando per l'anno in corso che è cominciato quello passato
					Anno = (Now.Year - 1).ToString.Trim()
				End If

				Try
					Sql = "SELECT * FROM Anni Where Descrizione Like '%" & Anno & "/%'"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Ritorno = ""
						If Rec.Eof Then
							Sql = "SELECT * FROM Anni Where Descrizione Like '%" & Anno & "'"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Not Rec.Eof Then
									Do Until Rec.Eof
										Ritorno &= Rec("idAnno").Value & ";" &
											Rec("Descrizione").Value & ";" &
											Rec("NomeSquadra").Value & ";" &
											"§"
										Rec.MoveNext()
									Loop
									Rec.Close()
								Else
									Ritorno = "ERROR: Nessun valore rilevato"
								End If
							End If
						Else
							Do Until Rec.Eof
								Ritorno &= Rec("idAnno").Value & ";" &
								Rec("Descrizione").Value & ";" &
								Rec("NomeSquadra").Value & ";" &
								"§"
								Rec.MoveNext()
							Loop
							Rec.Close()
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
	Public Function RitornaTipologie(Squadra As String) As String
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
					Sql = "SELECT * FROM [Generale].[dbo].[TipologiePartite] Order By Descrizione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " No tipologies found"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idTipologia").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

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
	Public Function RitornaRuoli(Squadra As String) As String
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
				Dim Sql As String = ""

				Try
					Sql = "SELECT * From Ruoli Order By idRuolo"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun ruolo rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idRuolo").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

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
	Public Function RitornaMaxAnno(Squadra As String) As String
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
				Dim idAnno As Integer

				Try
					Sql = "SELECT Max(idAnno) From Anni"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " No years found"
							idAnno = -1
						Else
							idAnno = Rec(0).Value
							Ritorno = (Rec(0).Value) + 1 & ";"
						End If
						Rec.Close()
					End If

					If idAnno > -1 Then
						Sql = "SELECT Descrizione From Anni Where idAnno=" & idAnno
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " No description found"
							Else
								Dim desc As String = Rec(0).Value
								If desc.Contains("/") Then
									Dim c() As String = desc.Split("/")
									desc = Val(c(0) + 1).ToString & "/" & Val(c(1) + 1).ToString
								End If

								Ritorno &= desc & ";"
							End If
							Rec.Close()
						End If
					Else
						Ritorno = StringaErrore & " Nessun anno rilevato"
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
	Public Function CreaNuovoAnno(Squadra As String, idAnno As String, descAnno As String, nomeSquadra As String, idAnnoAttuale As String,
								  idUtente As String, CreazioneTuttiIDati As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim lat As String = ""
				Dim lon As String = ""
				Dim ind As String = ""

				Try
					Sql = "Select * From Anni Where Upper(Trim(NomeSquadra))= '" & nomeSquadra.Trim.ToUpper & "'"
					Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							lat = "0"
							lon = "0"
							ind = "Sconosciuto"
						Else
							lat = "" & Rec("Lat").Value.ToString
							lon = "" & Rec("Lon").Value.ToString
							ind = "" & Rec("Indirizzo").Value.ToString
							nomeSquadra = "" & Rec("NomeSquadra").Value.ToString

							lat = lat.Replace(",", ".")
							lon = lon.Replace(",", ".")
						End If
					End If

					Sql = "Insert Into Anni Values (" &
						" " & idAnno & ", " &
						"'" & descAnno.Replace(";", "_").Replace("'", "''") & "', " &
						"'" & nomeSquadra.Replace(";", "_").Replace("'", "''") & "', " &
						" " & lat & ", " &
						" " & lon & ", " &
						"'" & ind & "', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"'', " &
						"0, " &
						"'', " &
						"'', " &
						"'N' " &
						")"
					Ritorno = EsegueSql(Conn, Sql, Connessione)

					If Ritorno = "*" Then
						' Creazione utenti
						Sql = "Insert Into UtentiMobile SELECT " & idAnno & " as idAnno, idUtente, Utente, Cognome, Nome, PassWord, " &
							"EMail, idCategoria, idTipologia From UtentiMobile Where idAnno=" & idAnnoAttuale & " And idUtente=" & idUtente
						Ritorno = EsegueSql(Conn, Sql, Connessione)

						If Ritorno <> "*" Then
							EliminaDatiNuovoAnnoDopoErrore(idAnno, Conn, Connessione)
						Else
							If CreazioneTuttiIDati = "S" Then
								' Creazione categorie
								Sql = "Insert Into Categorie SELECT " & idAnno & " as idAnno, idCategoria, Descrizione, Eliminato, " &
									"Ordinamento From Categorie Where Eliminato='N' And idAnno=" & idAnnoAttuale
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno <> "*" Then
									EliminaDatiNuovoAnnoDopoErrore(idAnno, Conn, Connessione)
								Else
									' Allenatori
									Sql = "Insert Into Allenatori " &
										"Select " & idAnno & " As idAnno, idAllenatore, Cognome, Nome, EMail, Telefono, Eliminato, idCategoria  From " &
										"Allenatori Where Eliminato='N' And idAnno=" & idAnnoAttuale
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno <> "*" Then
										EliminaDatiNuovoAnnoDopoErrore(idAnno, Conn, Connessione)
									Else
										' Dirigenti
										Sql = "Insert Into Dirigenti " &
											"SELECT " & idAnno & " as idAnno, idDirigente, Cognome, Nome, EMail, Telefono, Eliminato, idCategoria From " &
											"Dirigenti Where Eliminato='N' And idAnno=" & idAnnoAttuale
										Ritorno = EsegueSql(Conn, Sql, Connessione)
										If Ritorno <> "*" Then
											EliminaDatiNuovoAnnoDopoErrore(idAnno, Conn, Connessione)
										Else
											' Giocatori
											Sql = "Insert Into Giocatori " &
												"SELECT " & idAnno & " as idAnno, idGiocatore, idCategoria, idRuolo, Cognome, Nome, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, CodFiscale, Eliminato, CertScad, Maschio, " &
												"Telefono2, Citta, idTaglia, idCategoria2, Matricola, NumeroMaglia, idCategoria3 From Giocatori Where idAnno=" & idAnnoAttuale & " And Eliminato='N'"
											Ritorno = EsegueSql(Conn, Sql, Connessione)
											If Ritorno <> "*" Then
												EliminaDatiNuovoAnnoDopoErrore(idAnno, Conn, Connessione)
											Else
												' Arbitri
												Sql = "Insert Into Arbitri " &
												"SELECT " & idAnno & " As idAnno, idArbitro, Cognome, Nome, EMail, Telefono, Eliminato, idCategoria From Arbitri Where idAnno=" & idAnnoAttuale & " And Eliminato='N'"
												Ritorno = EsegueSql(Conn, Sql, Connessione)
												If Ritorno <> "*" Then
													EliminaDatiNuovoAnnoDopoErrore(idAnno, Conn, Connessione)
												End If
											End If
										End If
									End If
								End If
							End If
						End If
					End If

				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				' Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaAnnoAttualeUtente(Squadra As String, idUtente As String) As String
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
				Dim idAnno As Integer

				Try
					Sql = "SELECT * From AnnoAttualeUtenti Where idUtente=" & idUtente
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Rec.Close

							Sql = "SELECT Max(idAnno) From Anni"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									idAnno = 1
								Else
									idAnno = Rec(0).Value
									Ritorno = (Rec(0).Value) & ";"
								End If
							End If
						Else
							idAnno = Rec(1).Value
							Ritorno = (Rec(1).Value) & ";"
						End If
						Rec.Close()
					End If

					If idAnno > -1 Then
						Sql = "SELECT Descrizione, NomeSquadra, Lat, Lon From Anni Where idAnno=" & idAnno
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " No description found"
							Else
								Dim desc As String = Rec(0).Value
								Dim NomeSquadra As String = "" & Rec("NomeSquadra").Value

								Ritorno &= desc & ";" & NomeSquadra & ";" & Rec("Lat").Value & ";" & Rec("Lon").Value & ";"
							End If
							Rec.Close()
						End If
					Else
						Ritorno = StringaErrore & " Nessun anno rilevato"
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
	Public Function ImpostaAnnoAttualeUtente(Squadra As String, idAnno As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim Sql As String = ""

			Sql = "Begin transaction"
			Ritorno = EsegueSql(Conn, Sql, Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Ok As Boolean = True

				Try
					Sql = "Delete From AnnoAttualeUtenti Where idUtente=" & idUtente
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						Sql = "Insert Into AnnoAttualeUtenti Values (" & idUtente & ", " & idAnno & ")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

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
	Public Function ControllaEsistenzaDB(Squadra As String) As String
		Dim Ritorno As String = ""

		' Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Sorgenti\VB.Net\Miei\WEB\SSDCastelverdeCalcio\CVCalcio\DB\DB_Ponte_Di_Nona.mdb;Persist Security Info=False
		Dim Percorso As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Percorso = Mid(Percorso, Percorso.IndexOf("Data Source=") + 13, Percorso.Length)
		Percorso = Mid(Percorso, 1, Percorso.IndexOf(";"))

		If File.Exists(Percorso) Then
			Ritorno = "*"
		Else
			Ritorno = "ERROR: squadra non gestita"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaPartitaHTML(Squadra As String, idAnno As String, idPartita As String) As String
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
	Public Function RitornaSquadrePerSceltaIniziale() As String
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
				Dim Sql As String = ""
				Dim Altro As String = ""

				Try
					Sql = "SELECT SquadreAvversarie.idAvversario, SquadreAvversarie.idCampo, SquadreAvversarie.Descrizione, CampiAvversari.Descrizione As Campo, Indirizzo, Lat, Lon " &
						"FROM (SquadreAvversarie " &
						"Left Join CampiAvversari On SquadreAvversarie.idCampo=CampiAvversari.idCampo) " &
						"Left Join AvversariCoord On AvversariCoord.idAvversario=SquadreAvversarie.idAvversario " &
						"Where SquadreAvversarie.Eliminato='N' " & Altro & "Order By SquadreAvversarie.Descrizione"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " No avversaries found"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idAvversario").Value.ToString & ";" & Rec("idCampo").Value.ToString & ";" & Rec("Descrizione").Value.ToString.Trim & ";" & Rec("Campo").Value.ToString.Trim & ";" &
									Rec("Indirizzo").Value.ToString.Trim & ";" & Rec("Lat").Value.ToString & ";" & Rec("Lon").Value.ToString & ";§"

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
	Public Function RichiedeFirma(Squadra As String, CodSquadra As String, idUtente As String, Mail As String, Privacy As String, Mittente As String) As String
		' RichiedeFirma?Squadra= 0002_00160&idGiocatore=432&Genitore=1 
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), CodSquadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim c() As String = CodSquadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim m As New mail
				Dim Oggetto As String = "Richiesta Firma inCalcio"
				Dim Body As String = ""

				Dim gf As New GestioneFilesDirectory
				Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PercorsoSito.txt")

				If Percorso = "" Then
					Ritorno = StringaErrore & " Nessun percorso sito rilevato"
				Else
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "/" Then
						Percorso &= "/"
					End If

					Body &= "E' stata richiesta la firma dalla società " & Squadra.Replace("_", " ") & ".<br /><br />"
					Body &= "Per effettuare l'operazione eseguire il seguente link:<br /><br />"

					Body &= "<a href= """ & Percorso & "?firma=true&codSquadra=" & CodSquadra & "&id=" & idUtente & "&squadra=" & Squadra.Replace(" ", "_") & "&anno=" & Anno & "&genitore=-1&privacy=" & Privacy & "&tipoUtente=2"">"
					Body &= "Click per firmare"
					Body &= "</a>"

					Ritorno = m.SendEmail(Squadra, Mittente, Oggetto, Body, Mail, {""})
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RinominaImmagini() As String
		Dim gf As New GestioneFilesDirectory
		Dim path As String = ""
		Dim quanteConversioni As Integer = 0

		path = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
		path = path.Replace(vbCrLf, "")
		If Strings.Right(path, 1) <> "\" Then
			path &= "\"
		End If

		gf.ScansionaDirectorySingola(path)
		Dim Filetti() As String = gf.RitornaFilesRilevati
		Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati
		Dim cr As New CriptaFiles

		For i As Integer = 1 To qFiletti
			Dim NomeOrigine As String = Filetti(i)
			Dim Controllo As String = NomeOrigine.Replace(path, "")
			If Controllo.Contains("\") Then
				Dim Estensione As String = gf.TornaEstensioneFileDaPath(Filetti(i))
				Dim NomeDestinazione As String = NomeOrigine
				NomeDestinazione = NomeDestinazione.Replace(Estensione, "") & ".kgb"

				If Not NomeOrigine.ToUpper.Contains("\APPOGGIO\") And Not NomeOrigine.ToUpper.Contains("\ICONE\") And (Estensione.ToUpper = ".JPG" Or Estensione.ToUpper = ".PNG") Then
					If Not File.Exists(NomeDestinazione) Then
						cr.EncryptFile(CryptPasswordString, NomeOrigine, NomeDestinazione)
						File.Delete(NomeOrigine)

						quanteConversioni += 1
					End If
				End If
			End If
		Next

		path = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		path = path.Replace(vbCrLf, "")
		Dim p() As String = path.Split(";")
		path = p(0)
		If Strings.Right(path, 1) <> "\" Then
			path &= "\"
		End If

		gf.ScansionaDirectorySingola(path)
		Filetti = gf.RitornaFilesRilevati
		qFiletti = gf.RitornaQuantiFilesRilevati

		For i As Integer = 1 To qFiletti
			Dim NomeOrigine As String = Filetti(i)
			Dim Controllo As String = NomeOrigine.Replace(path, "")
			If Controllo.Contains("\") Then
				Dim Estensione As String = gf.TornaEstensioneFileDaPath(Filetti(i))
				Dim NomeDestinazione As String = NomeOrigine
				NomeDestinazione = NomeDestinazione.Replace(Estensione, "") & ".kgb"

				If Not NomeOrigine.ToUpper.Contains("\APPOGGIO\") And Not NomeOrigine.ToUpper.Contains("\ICONE\") And (Estensione.ToUpper = ".JPG" Or Estensione.ToUpper = ".PNG") Then
					If Not File.Exists(NomeDestinazione) Then
						cr.EncryptFile(CryptPasswordString, NomeOrigine, NomeDestinazione)
						File.Delete(NomeOrigine)

						quanteConversioni += 1
					End If
				End If
			End If
		Next

		Return "Immagini convertite: " & quanteConversioni
	End Function

	<WebMethod()>
	Public Function PulisceCartellaTemp() As String
		Dim Ritorno As String = ""

		Ritorno = PulisceCartellaTemporanea()

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ConverteImmagine(CodiceSquadra As String, NomeSquadra As String, MultimediaOAllegati As String, Immagine As String) As String
		NomeSquadra = NomeSquadra.Replace(" ", "_")

		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory
		Dim path As String = ""
		Dim pathImmagine As String = ""
		Dim pathDestinazione As String = ""
		Dim fileOrigine As String = ""
		Dim fileDestinazione As String = ""
		Dim urlIniziale As String = ""
		Dim chiave As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
		Dim Datella As String = ""
		For I As Integer = 1 To 10
			Dim codice As Integer = RitornaValoreRandom(chiave.Length - 1) + 1
			Datella &= Mid(chiave, codice, 1)
		Next
		Dim nn As String = gf.TornaNomeFileDaPath(Immagine)
		If nn.Contains(".") Then
			nn = Mid(nn, 1, nn.IndexOf("."))
		End If
		Datella &= "_" & nn
		Dim Ok As Boolean = True
		Dim cr As New CriptaFiles

		' http://192.168.0.227:92/MultiMedia/Morti_De_Sonno_Fc/Giocatori/1_542.jpg
		If MultimediaOAllegati = "M" Then
			path = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
			path = path.Replace(vbCrLf, "")
			If Strings.Right(path, 1) <> "\" Then
				path &= "\"
			End If
			Dim a As Integer = Immagine.ToUpper.IndexOf(NomeSquadra.ToUpper)
			If a = -1 Then
				Ok = False
			Else
				urlIniziale = Mid(Immagine, 1, a)
				pathImmagine = Mid(Immagine, a + 1, Immagine.Length)
			End If
		Else
			path = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
			path = path.Replace(vbCrLf, "")
			Dim p() As String = path.Split(";")
			path = p(0)
			If Strings.Right(path, 1) <> "\" Then
				path &= "\"
			End If
			Dim a As Integer = Immagine.ToUpper.IndexOf(CodiceSquadra.ToUpper)
			If a = -1 Then
				Ok = False
			Else
				urlIniziale = Mid(Immagine, 1, a)
				pathImmagine = Mid(Immagine, a + 1, Immagine.Length)
			End If
		End If

		If Ok Then
			pathDestinazione = path & "Appoggio\"
			fileOrigine = path & pathImmagine
			fileDestinazione = pathDestinazione & Datella & ".jpg"

			fileOrigine = fileOrigine.Replace("%2F", "\")
			fileDestinazione = fileDestinazione.Replace("%2F", "\")

			If Immagine <> "" Then
				If File.Exists(fileOrigine) Then
					gf.CreaDirectoryDaPercorso(fileDestinazione)
					cr.DecryptFile(CryptPasswordString, fileOrigine, fileDestinazione)

					' File.Copy(fileOrigine, fileDestinazione)

					Ritorno = urlIniziale & "Appoggio/" & Datella & ".jpg"
					quanteConversioni += 1
					If quanteConversioni > 50 Then
						PulisceCartellaTemporanea()
						quanteConversioni = 0
					End If
					'Dim t As New Timer With {.Interval = 10000}
					't.Tag = DateTime.Now
					'AddHandler t.Tick, Sub(sender, e) MyTickHandler(t, fileDestinazione)
					't.Start()
				Else
					Ritorno = StringaErrore & " Immagine non esistente"
				End If
			Else
				Ritorno = StringaErrore & " Nessuna immagine passata"
			End If
		Else
			Ritorno = StringaErrore & " Errore nel decodificare la stringa"
		End If

		Return Ritorno
	End Function

	'Private Sub MyTickHandler(t As Timer, ritorno As String)
	'	File.Delete(ritorno)
	'	t.Stop()
	'	t.Dispose()
	'End Sub

	<WebMethod()>
	Public Function RitornaDatiContatti(Squadra As String, idCategoria As String) As String
		' RichiedeFirma?Squadra= 0002_00160&idGiocatore=432&Genitore=1 
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

				Sql = "Select * From (Select idGiocatore, DataDiNascita, Cognome, Nome, Maggiorenne, EMail As Dettaglio1, Telefono As Dettaglio2, '' As Dettaglio3, '' As Dettaglio4, '' As Dettaglio5, '' As Dettaglio6 From Giocatori " &
					"Where Eliminato = 'N' And CharIndex('" & idCategoria & "-', Categorie) > 0 And Maggiorenne = 'S' " &
					"Union All " &
					"Select A.idGiocatore, DataDiNascita, Cognome, Nome, A.Maggiorenne, Genitore1 As Dettaglio1, MailGenitore1 As Dettaglio2, TelefonoGenitore1 As Dettaglio3, " &
					"Genitore2 As Dettaglio4, MailGenitore2 As Dettaglio5, TelefonoGenitore2 As Dettaglio6 From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Where A.Eliminato = 'N' And CharIndex('" & idCategoria & "-', A.Categorie) > 0 And A.Maggiorenne = 'N' " &
					") A Order By Cognome, Nome"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Ritorno = ""
					Do Until Rec.Eof
						Ritorno &= Rec("Cognome").Value & ";"
						Ritorno &= Rec("Nome").Value & ";"
						Ritorno &= Rec("Maggiorenne").Value & ";"
						Ritorno &= Rec("Dettaglio1").Value & ";"
						Ritorno &= Rec("Dettaglio2").Value & ";"
						Ritorno &= Rec("Dettaglio3").Value & ";"
						Ritorno &= Rec("Dettaglio4").Value & ";"
						Ritorno &= Rec("Dettaglio5").Value & ";"
						Ritorno &= Rec("Dettaglio6").Value & ";"
						Ritorno &= Rec("idGiocatore").Value & ";"
						Ritorno &= Rec("DataDiNascita").Value & ";"
						Ritorno &= "§"

						Rec.MoveNext
					Loop
					Rec.Close
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StampaHTML(fileToPrint As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim cont As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim pp() As String = cont.Split(";")
		Dim path As String = pp(2).Replace(vbCrLf, "")
		path = path.Replace("Multimedia", "Allegati")
		If Strings.Right(path, 1) <> "/" Then
			path &= "/"
		End If
		Dim pathFisico As String = pp(0).Trim
		If Strings.Right(pathFisico, 1) <> "\" Then
			pathFisico &= "\"
		End If
		fileToPrint = fileToPrint.Replace(path, pathFisico)
		fileToPrint = fileToPrint.Replace("/", "\")

		Using printProcess As Process = New Process()
			Dim systemPath As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
			printProcess.StartInfo.FileName = systemPath + "\rundll32.exe"
			printProcess.StartInfo.Arguments = systemPath + "\mshtml.dll,PrintHTML """ & fileToPrint & """"
			printProcess.Start()
		End Using

		Return fileToPrint
	End Function
End Class