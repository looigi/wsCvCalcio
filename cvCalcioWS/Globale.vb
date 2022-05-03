Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Timers
Imports System.Windows.Forms

Module Globale
	Public effettuaLog As Boolean = True
	Public effettuaLogMail As Boolean = True

	Public nomeFileLogMail As String = ""

	Public quanteConversioni As Integer = 0

	Private Structure Sostituz
		Dim Minuto As Integer
		Dim Tempo As Integer
		Dim idEntrato As Integer
		Dim idUscito As Integer
		Dim NominativoEntrato As String
		Dim NominativoUscito As String
		Dim SoprannomeEntrato As String
		Dim SoprannomeUscito As String
	End Structure

	Private Structure AmmoEspu
		Dim idGiocatore As Integer
		Dim Minuto As Integer
		Dim Tempo As Integer
	End Structure

	Public Structure FormatoByte
		Dim Occupazione As Double
		Dim Cosa As String
	End Structure

	Public Structure strutturaMail
		Dim Squadra As String
		Dim Mittente As String
		Dim Oggetto As String
		Dim newBody As String
		Dim Ricevente As String
		Dim Allegato() As String
		Dim AllegatoOMultimedia As String
		Dim NuovaSocieta As String
	End Structure
	Public listaMails As New List(Of strutturaMail)
	Public timerMails As Timers.Timer = Nothing
	Public pathMail As String = ""

	Public Const ErroreConnessioneNonValida As String = "ERRORE: Stringa di connessione non valida"
	Public Const ErroreConnessioneDBNonValida As String = "ERRORE: Connessione al db non valida"
	Public Percorso As String
	' Public PercorsoSitoCV As String = "C:\GestioneCampionato\CalcioImages\" ' "C:\inetpub\wwwroot\CVCalcio\App_Themes\Standard\Images\"
	' Public PercorsoSitoURLImmagini As String = "http://loppa.duckdns.org:90/MultiMedia/" ' "http://looigi.no-ip.biz:90/CvCalcio/App_Themes/Standard/Images/"
	Public StringaErrore As String = "ERROR: "
	Public RigaPari As Boolean = False
	Public CryptPasswordString As String = "WPippoBaudo227!"
	Public stringaWidgets As String = "1-1-1-1-1"
	Public TipoDB As String = "MariaDB"
	Public TipoPATH As String = "SQLSERVER"
	Public TipoPDF As String = "WINDOWS"
	'Public mdb(5) As clsMariaDB

	'Public CodicePartitaPerImmagini As String = ""

	Public Function ConvertePath(Path As String) As String
		Dim NomeFile As String = Path

		If TipoPATH <> "SQLSERVER" Then
			NomeFile = NomeFile.Replace("\", "/")
			NomeFile = NomeFile.Replace("/\", "/")
			Dim inizio As String = Mid(NomeFile, 1, 10)
			Dim resto As String = Mid(NomeFile, 10, NomeFile.Length)
			resto = resto.Replace("//", "/")
			NomeFile = inizio & resto
		End If

		Return NomeFile
	End Function

	Public Function LeggeTipoDB() As String
		'If TipoDB = "" Then
		'	Dim gf As New GestioneFilesDirectory

		'	Return gf.LeggeFileIntero(MP & "\Impostazioni\TipoDB.txt")
		'Else
		Return TipoDB
		'End If
	End Function

	Public Function SistemaNumero(Numero As String) As String
		If Numero = "" Then
			Return "Null"
		Else
			Return Numero.Replace(",", ".")
		End If
	End Function

	Public Function AggiungeRigoriEGoal(NomeLista As List(Of String), Rec As Object) As List(Of String)
		Dim ListaNomi As New List(Of String)
		Dim posi As Integer = 0
		Dim Ok As Boolean = False

		ListaNomi.AddRange(NomeLista)
		For Each s As String In ListaNomi
			Dim n As Integer = Val(Mid(s, 1, s.IndexOf("-")))
			If n = Rec(4).Value Then
				Dim cc() As String = s.Split("-")
				ListaNomi.Item(posi) = Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & (Rec(3).Value + cc(2))
				Ok = True
				Exit For
			End If
			posi += 1
		Next

		If Not Ok Then
			ListaNomi.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
		End If

		Return ListaNomi
	End Function

	Public Sub EliminaDatiNuovoAnnoDopoErrore(MP As String, idAnno As String, Conn As Object, Connessione As String)
		Dim Ritorno As String
		Dim Sql As String

		Sql = "delete from Anni Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(MP, Sql, Connessione)

		Sql = "delete from UtentiMobile Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(MP, Sql, Connessione)

		Sql = "delete from Categorie Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(MP, Sql, Connessione)

		Sql = "delete from Allenatori Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(MP, Sql, Connessione)

		Sql = "delete from Dirigenti Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(MP, Sql, Connessione)

		Sql = "delete from Giocatori Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(MP, Sql, Connessione)

		Sql = "delete from Arbitri Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
	End Sub

	Public Function RitornaMailDopoRichiesta(MP As String, Squadra As String, Utente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(MP, "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Sql = "SELECT Utenti.idAnno, idUtente, Utente, Cognome, Nome, " &
						"Password, EMail, idCategoria, Utenti.idTipologia, Utenti.idSquadra, Descrizione As Squadra " &
						"FROM Utenti Left Join Squadre On Utenti.idSquadra = Squadre.idSquadra " &
						"Where Upper(Utente)='" & Utente.ToUpper.Replace("'", "''") & "'" ' And idAnno=" & idAnno
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec.ToString
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun utente rilevato"
					Else
						If Rec("EMail").Value = "" And Rec("Utente").Value = "" Then
							Ritorno = StringaErrore & " Nessuna mail rilevata"
						Else
							Dim idUtente As Integer = Rec("idUtente").Value
							Dim EMail As String = Rec("EMail").Value
							If EMail = "" Then
								EMail = Rec("Utente").Value
							End If
							Dim pass As String = generaPassRandom()
							Dim nuovaPass() = pass.Split(";")

							Try
								Sql = "Update Utenti Set Password='" & nuovaPass(1).Replace("'", "''") & "', PasswordScaduta=1 " &
									"Where idUtente=" & idUtente
								Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
								If Not Ritorno.Contains(StringaErrore) Then
									Dim m As New mail
									Dim Oggetto As String = "Reset password inCalcio"
									Dim Body As String = ""
									Body &= "La Sua password relativa al sito inCalcio è stata modificata dietro sua richiesta. <br /><br />"
									Body &= "La nuova password valida per il solo primo accesso è: " & nuovaPass(0) & "<br /><br />"
									Dim ChiScrive As String = "notifiche@incalcio.cloud"

									Ritorno = m.SendEmail(MP, "", "", Oggetto, Body, EMail, {""})
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
							End Try
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function LeggeImpostazioniDiBase(Percorso As String, Squadra As String) As String
		Dim Connessione As String = ""

		' Impostazioni di base
		Dim ListaConnessioni As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

		If ListaConnessioni.Count <> 0 Then
			Dim NomeStringa As String = "SQLConnectionStringLOCALEMD"

			' Get the collection elements. 
			For Each Connessioni As ConnectionStringSettings In ListaConnessioni
				Dim Nome As String = Connessioni.Name
				Dim Provider As String = Connessioni.ProviderName
				Dim connectionString As String = Connessioni.ConnectionString

				If TipoDB = "SQLSERVER" Then
					If Nome = "SQLConnectionStringLOCALESS" Then
						Connessione = "Provider=" & Provider & ";" & connectionString
						Connessione = Replace(Connessione, "*^*^*", Percorso & "\")
						If Squadra <> "" Then
							If Squadra = "DBVUOTO" Then
								Connessione = Connessione.Replace("***NOME_DB***", "DBVuoto")
							Else
								Connessione = Connessione.Replace("***NOME_DB***", Squadra)
							End If
						Else
							Connessione = Connessione.Replace("***NOME_DB***", "Generale")
						End If
						Exit For
					End If
				Else
					If Nome = NomeStringa Then
						Connessione = connectionString
						Connessione = Replace(Connessione, "*^*^*", Percorso & "\")
						If Squadra <> "" Then
							If Squadra = "DBVUOTO" Then
								Connessione = Connessione.Replace("***NOME_DB***", "DBVuoto")
							Else
								Connessione = Connessione.Replace("***NOME_DB***", Squadra)
							End If
						Else
							Connessione = Connessione.Replace("***NOME_DB***", "Generale")
						End If
						Exit For
					End If
				End If
			Next
		End If

		Return Connessione
	End Function

	Public Function RitornaMultimediaPerTipologia(MP As String, Squadra As String, idAnno As String, id As String, Tipologia As String) As String
		' PercorsoSitoCV = "D:\Looigi\VB.Net\Miei\WEB\SSDCastelverdeCalcio\CVCalcio\App_Themes\Standard\Images\"
		Dim gf As New GestioneFilesDirectory
		Dim Righe As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")
		Dim Campi() As String = Righe.Split(";")
		Campi(0) = Campi(0).Replace(vbCrLf, "")
		If Strings.Right(Campi(0), 1) <> "\" Then
			Campi(0) = Campi(0) & "\"
		End If
		Campi(2) = Campi(2).Replace(vbCrLf, "")
		If Strings.Right(Campi(2), 1) <> "/" Then
			Campi(2) = Campi(2) & "/"
		End If
		Campi(2) = Campi(2).Replace("Multimedia", "Allegati")

		Dim Ritorno As String = ""
		Dim Ok As Boolean = True
		Dim Percorso As String = Campi(0) & Squadra & "\" & Tipologia & "\Anno" & idAnno & "\"
		Percorso = Percorso.Replace("\\", "\")
		Dim IndirizzoURL As String = Campi(2) & Squadra & "/" & Tipologia & "/Anno" & idAnno & "/"
		IndirizzoURL = IndirizzoURL.Replace("//", "/")
		Dim Codice As String

		Select Case Tipologia
			Case "Partite"
				Codice = id.ToString
				For i As Integer = Codice.Length + 1 To 5
					Codice = "0" & Codice
				Next
			Case Else
				Codice = idAnno.ToString & "_" & id.ToString
		End Select
		Percorso &= Codice
		IndirizzoURL &= Codice & "/"
		gf.CreaDirectoryDaPercorso(Percorso & "\")
		gf.ScansionaDirectorySingola(Percorso)
		Dim Filetti() As String = gf.RitornaFilesRilevati
		Dim qFiletti As String = gf.RitornaQuantiFilesRilevati
		Dim Estensioni() As String = {".JPG", ".JPEG", ".BMP", ".PNG", ".GIF"}

		For i As Integer = 1 To qFiletti
			Dim Ok2 As Boolean = False
			For Each e As String In Estensioni
				If Filetti(i).ToUpper.Trim.IndexOf(e) > -1 Then
					Ok2 = True
					Exit For
				End If
			Next
			If Ok2 Then
				Dim Barra As String = "\"

				If TipoDB <> "SQLSERVER" Then
					Barra = "/"
				End If

				Dim Dimensioni As Long = FileLen(Filetti(i))
				Dim DataUltimaModifica As String = gf.TornaDataDiUltimaModifica(Filetti(i))
				Dim NomeUrl As String = Filetti(i).Replace(Percorso & Barra, "")
				Dim NomeFile As String = gf.TornaNomeFileDaPath(NomeUrl)

				Ritorno &= NomeUrl & ";" & NomeFile & ";" & Dimensioni.ToString & ";" & DataUltimaModifica & ";" & Codice & "§"
			End If
		Next

		If Ritorno = "" Then
			Ritorno = StringaErrore & " Nessun file rilevato"
		End If

		Return Ritorno
	End Function

	Public Function DecriptaImmagine(MP As String, Nome As String) As String
		Dim Ritorno As String = ""
		'Dim gf As New GestioneFilesDirectory
		'Dim Barra As String = "\"
		'Dim ControBarra As String = "/"

		'If TipoDB <> "SQLSERVER" Then Barra = "/"

		'Dim tutto As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")
		'Dim campi() As String = tutto.Split(";")
		'Dim pathFisico As String = campi(0).Replace(vbCrLf, "")
		'Dim pathUrl As String = campi(5).Replace(vbCrLf, "")

		'If Strings.Right(pathFisico, 1) <> Barra Then pathFisico &= Barra
		'If Strings.Right(pathUrl, 1) <> ControBarra Then pathUrl &= ControBarra
		''If TipoDB = "SQLSERVER" Then
		''	pathFisico = pathFisico.Replace("Allegati", "CalcioImages")
		''Else
		''	pathUrl = pathUrl.Replace("Multimedia/allegati", "Multimedia/multimedia")
		''End If

		'Dim pathLetturaFile1 As String = Nome.Replace(pathUrl, "")

		'pathLetturaFile1 = pathLetturaFile1.Replace(ControBarra, Barra)
		''pathLetturaFile1 = pathFisico & pathLetturaFile1
		'pathLetturaFile1 = pathLetturaFile1.Replace(Barra & Barra, Barra)
		'pathLetturaFile1 = pathLetturaFile1.Replace(" ", "_")

		'Dim pathAppoggio As String = pathFisico & "Appoggio"
		'Dim stringaRandom As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
		'Dim r As String = ""
		'For i As Integer = 1 To 5
		'	Dim p As String = RitornaValoreRandom(stringaRandom.Length - 1) + 1
		'	r &= Mid(stringaRandom, p, 1)
		'Next
		'Dim NomeFile As String = r & "_" & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute & "00") & Format(Now.Second, "00") & ".jpg"
		'Dim pathScritturaFile1 As String = pathAppoggio & Barra & NomeFile
		'Dim pathUrl1 As String = pathUrl & "multimedia/Appoggio" & Barra & NomeFile

		'Dim PathBaseImmScon As String = pathUrl & "Sconosciuto.png"

		'Dim c As New CriptaFiles

		'If TipoDB <> "SQLSERVER" Then
		'	pathLetturaFile1 = pathLetturaFile1.Replace("Multimedia/allegati", "Multimedia/multimedia")
		'	pathScritturaFile1 = pathScritturaFile1.Replace("Multimedia/allegati", "Multimedia/multimedia")
		'End If

		''Return pathLetturaFile1 & " - " & pathScritturaFile1 & " - " & pathUrl1 & PathBaseImmScon

		'If ControllaEsistenzaFile(pathLetturaFile1) Then
		'	c.DecryptFile(CryptPasswordString, pathLetturaFile1, pathScritturaFile1)
		'	If ControllaEsistenzaFile(pathScritturaFile1) Then
		'		Ritorno = pathUrl1
		'	Else
		'		Ritorno = PathBaseImmScon
		'	End If
		'Else
		'	Ritorno = PathBaseImmScon
		'End If

		Return Ritorno
	End Function

	Public Function CreaHtmlPartita(MP As String, Squadra As String, Conn As Object, Connessione As String, idAnno As String, idPartita As String) As String
		Dim Sql As String
		Dim Rec As Object
		Dim Rec2 As Object
		Dim Ok As Boolean = True
		Dim Pagina As StringBuilder = New StringBuilder
		Dim gf As New GestioneFilesDirectory
		Dim Barra As String = "\"

		If TipoDB <> "SQLSERVER" Then
			Barra = "/"
		End If

		Dim paths As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")
		Dim P() As String = paths.Split(";")
		P(0) = P(0).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(0), 1) <> Barra Then
			P(0) &= Barra
		End If
		Dim pathAllegati As String = P(0).Replace(vbCrLf, "")

		P(1) = P(1).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(1), 1) <> Barra Then
			P(1) &= Barra
		End If
		Dim pathLogG As String = P(1).Replace(vbCrLf, "")

		P(2) = P(2).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(2), 1) <> Barra Then
			P(2) &= Barra
		End If
		Dim pathMultimedia As String = P(0).Replace(vbCrLf, "")

		P(4) = P(4).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(4), 1) <> Barra Then
			P(4) &= Barra
		End If
		Dim urlAllegati As String = P(4)

		Dim PathBaseImmagini As String = pathMultimedia
		Dim PathBaseMultimedia As String = pathMultimedia.Replace("Allegati", "Multimedia")
		Dim PathBaseImmScon As String = pathMultimedia & "Sconosciuto.png"
		Dim PathLog As String = pathLogG & "Pdf.txt"
		Dim Ritorno As String = "*"

		Dim Filone As String = gf.LeggeFileIntero(MP & "\Scheletri\base_partita.txt")
		Dim sIdPartita As String = idPartita.Trim
		For i As Integer = sIdPartita.Length - 1 To 3
			sIdPartita = "0" & sIdPartita
		Next
		Dim NomeFileFinale As String = pathAllegati & Squadra & "\Partite\Anno" & idAnno & "\" & sIdPartita & "\" & idPartita & ".html"
		Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Partite\Anno" & idAnno & "\" & sIdPartita & "\" & idPartita & ".pdf"
		Dim PathPerMultimedia As String = pathAllegati & Squadra & "\Partite\Anno" & idAnno & "\" & sIdPartita & "\"

		Dim altezzaReport As Integer = 350 ' Intestazione
		Dim altezzaConvocati As Integer = 0
		Dim altezzaEventi As Integer = 0

		gf.CreaDirectoryDaPercorso(NomeFileFinale)
		gf.EliminaFileFisico(NomeFileFinale)
		gf.EliminaFileFisico(NomeFileFinalePDF)

		'gf.ApreFileDiTestoPerScrittura(NomeFileFinalePDF)
		'gf.ScriveTestoSuFileAperto("ppp")
		'gf.ChiudeFileDiTestoDopoScrittura()

		If TipoPATH <> "SQLSERVER" Then
			PathPerMultimedia = PathPerMultimedia.Replace("\", "/")
			PathPerMultimedia = PathPerMultimedia.Replace("//", "/")

			PathBaseImmagini = PathBaseImmagini.Replace("Multimedia/allegati", "Multimedia/multimedia")
		End If

		'Return PathBaseImmagini

		gf.ScansionaDirectorySingola(PathPerMultimedia)
		Dim Multimedia() As String = gf.RitornaFilesRilevati
		Dim qMultimedia As Integer = gf.RitornaQuantiFilesRilevati
		Dim mmu As String = ""
		For i As Integer = 1 To qMultimedia
			Dim este As String = gf.TornaEstensioneFileDaPath(Multimedia(i)).ToUpper.Trim.Replace(".", "")
			Dim PathEsternoMM As String = Multimedia(i).Replace(pathAllegati, pathMultimedia) ' .Replace("Multimedia", "Allegati").Replace("\", "/")
			Dim Nome As String = gf.TornaNomeFileDaPath(Multimedia(i))
			Dim e As String = gf.TornaEstensioneFileDaPath(Nome)

			Nome = Nome.Replace(e, "")
			If Nome.Length > 14 Then
				Nome = Mid(Nome, 1, 6) & "..." & Mid(Nome, Nome.Length - 6, 6)
			End If

			If este = "JPG" Or este = "JPEG" Or este = "JFIF" Or este = "BMP" Or este = "GIF" Or este = "PNG" Then
				' Immagine
				mmu &= "<div class=""multimedia""><a href=""" & PathEsternoMM & """ target=""_blank""><img src=""" & PathEsternoMM & """ width=""100%"" height=""100%""><br /><span class=""testo nero"" style=""font-size: 10px; white-space: nowrap;"">" & Nome & "</span></a></div>" & vbCrLf
			Else
				If este = "MP4" Or este = "WMV" Or este = "MPG" Then
					' Filmato
					mmu &= "<div class=""multimedia""><a href=""" & PathEsternoMM & """ target=""_blank""><img src=""" & PathBaseImmagini & "/video.png"" width=""100%"" height=""100%""><br /><span class=""testo nero"" style=""font-size: 10px; white-space: nowrap;"">" & Nome & "</span></a></div>" & vbCrLf
				End If
			End If
		Next
		If mmu <> "" Then
			Filone = Filone.Replace("***MMVISIBILE***", "block")
			Filone = Filone.Replace("***MULTIMEDIA***", mmu)
		Else
			Filone = Filone.Replace("***MMVISIBILE***", "none")
		End If

		If qMultimedia > 0 Then
			Dim righe As Integer = CInt(qMultimedia / 8)
			If righe = 0 Then righe = 1

			altezzaReport += (righe * 110) + 70
		End If
		' Multimedia

		' Tempi partita
		altezzaReport += 50

		Filone = Filone.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")
		' Filone = Filone.Replace("***SFONDO***", "")

		Dim Sostituzioni As New List(Of Sostituz)

		Sql = "Select B.Minuto, B.Tempo, C.idGiocatore As idEntrato, C.Cognome As CognomeE, C.Soprannome As SoprannomeE, C.Nome As NomeE, D.idGiocatore As idUscito, D.Cognome As CognomeU, D.Soprannome As SoprannomeU, D.Nome As NomeU From Convocati A " &
			"Left Join PartiteSostituzioni B On A.idPartita = B.idPartita And A.idGiocatore = B.idSostituito " &
			"Left Join Giocatori C On C.idGiocatore = B.idEntrante " &
			"Left Join Giocatori D On D.idGiocatore = B.idSostituito " &
			"Where A.idPartita = " & idPartita & " And B.idPartita Is Not Null"
		Rec = Conn.LeggeQuery(MP, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ok = False
			Ritorno = "Problemi lettura generale"
		Else
			Do Until Rec.Eof()
				Dim s As New Sostituz
				s.idEntrato = Rec("idEntrato").Value
				s.idUscito = Rec("idUscito").Value
				s.Minuto = Rec("Minuto").Value
				s.Tempo = Rec("Tempo").Value
				s.NominativoEntrato = Rec("CognomeE").Value & " " & Rec("NomeE").Value
				s.NominativoUscito = Rec("CognomeU").Value & " " & Rec("NomeU").Value
				s.SoprannomeEntrato = Rec("SoprannomeE").Value
				s.SoprannomeUscito = Rec("SoprannomeU").Value

				Sostituzioni.Add(s)

				Rec.MoveNext()
			Loop
			Rec.Close()
		End If

		Dim InFormazione As New List(Of Integer)

		Sql = "Select * From Titolari " &
			"Where idPartita=" & idPartita
		Rec2 = Conn.LeggeQuery(MP, Sql, Connessione)
		If TypeOf (Rec2) Is String Then
		Else
			Do Until Rec2.Eof()
				InFormazione.Add(Rec2("idGiocatore").Value)
				Rec2.MoveNext
			Loop
		End If
		Rec2.Close()

		Sql = "SELECT Partite.idPartita, Partite.idCategoria, Partite.idAvversario, Partite.idTipologia, Partite.idCampo, " &
			"Partite.idUnioneCalendario, Partite.DataOra, Partite.Giocata, Partite.OraConv, Risultati.Risultato, Risultati.Note, " &
			"RisultatiAggiuntivi.RisGiochetti, RisultatiAggiuntivi.GoalAvvPrimoTempo, RisultatiAggiuntivi.GoalAvvSecondoTempo, " &
			"RisultatiAggiuntivi.GoalAvvTerzoTempo, SquadreAvversarie.Descrizione AS Avversario, CampiAvversari.Descrizione AS CampoA, " &
			"TipologiePartite.Descrizione AS Tipologia, " & IIf(TipoDB = "SQLSERVER", "Allenatori.Cognome+' '+Allenatori.Nome", "Concat(Allenatori.Cognome,' ',Allenatori.Nome)") & " AS Allenatore, " &
			" " & IIf(TipoDB = "SQLSERVER", "Categorie.AnnoCategoria + '-' + Categorie.Descrizione", "Concat(Categorie.AnnoCategoria, '-', Categorie.Descrizione)") & " As Categoria, " &
			"CampiAvversari.Indirizzo as CampoIndirizzo, Partite.Casa, Allenatori.idAllenatore, CampiEsterni.Descrizione As CampoEsterno, " &
			"RisultatiAggiuntivi.Tempo1Tempo, RisultatiAggiuntivi.Tempo2Tempo, RisultatiAggiuntivi.Tempo3Tempo, " &
			"CoordinatePartite.Lat, CoordinatePartite.Lon, TempiGoalAvversari.TempiPrimoTempo, TempiGoalAvversari.TempiSecondoTempo, TempiGoalAvversari.TempiTerzoTempo, " &
			"MeteoPartite.Tempo, MeteoPartite.Gradi, MeteoPartite.Umidita, MeteoPartite.Pressione, MeteoPartite.Icona, ArbitriPartite.idArbitro, " &
			" " & IIf(TipoDB = "SQLSERVER", "Arbitri.Cognome + ' ' + Arbitri.Nome", "Concat(Arbitri.Cognome, ' ', Arbitri.Nome)") & " As Arbitro, " &
			"Partite.RisultatoATempi, Partite.DataOraAppuntamento, Partite.LuogoAppuntamento, Partite.MezzoTrasporto, Categorie.AnticipoConvocazione, Anni.Indirizzo, Anni.Lat, Anni.Lon, " &
			"Anni.CampoSquadra, Anni.NomePolisportiva, Partite.ShootOut, Partite.Tempi, Partite.PartitaConRigori, " & IIf(TipoDB = "SQLSERVER", "IsNull(PartiteCapitani.idCapitano,0)", "Coalesce(PartiteCapitani.idCapitano,0)") & " As idCapitano " &
			"FROM Partite LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita " &
			"LEFT JOIN RisultatiAggiuntivi ON Partite.idPartita = RisultatiAggiuntivi.idPartita " &
			"LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario " &
			"LEFT JOIN [Generale].[dbo].[TipologiePartite] ON Partite.idTipologia = TipologiePartite.idTipologia " &
			"LEFT JOIN Allenatori ON Partite.idAnno = Allenatori.idAnno And Partite.idAllenatore = Allenatori.idAllenatore " &
			"LEFT JOIN CampiAvversari ON SquadreAvversarie.idCampo = CampiAvversari.idCampo " &
			"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita " &
			"LEFT JOIN Categorie ON Partite.idCategoria = Categorie.idCategoria And Categorie.idAnno = Partite.idAnno " &
			"LEFT JOIN CoordinatePartite On Partite.idPartita = CoordinatePartite.idPartita " &
			"LEFT JOIN MeteoPartite On Partite.idPartita = MeteoPartite.idPartita " &
			"LEFT JOIN TempiGoalAvversari On Partite.idPartita = TempiGoalAvversari.idPartita " &
			"LEFT JOIN ArbitriPartite On Partite.idPartita = ArbitriPartite.idPartita And ArbitriPartite.idAnno = Partite.idAnno " &
			"LEFT JOIN Arbitri On ArbitriPartite.idArbitro=Arbitri.idArbitro And ArbitriPartite.idAnno=Arbitri.idAnno " &
			"LEFT JOIN Anni On Partite.idAnno = Anni.idAnno " &
			"LEFT JOIN PartiteCapitani On Partite.idPartita = PartiteCapitani.idPartita " &
			"WHERE Partite.idPartita=" & idPartita & " And Partite.idAnno=" & idAnno
		Rec = Conn.LeggeQuery(MP, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ok = False
			Ritorno = "Problemi lettura generale"
		Else
			Dim idCapitano As Integer = -1

			If Not Rec.Eof() Then
				'If Not Rec("idCapitano").Value Is DBNull.Value Then
				idCapitano = Rec("idCapitano").Value
				'End If

				Dim immMeteo As String = "<img src=""" & Rec("Icona").Value & """ style=""width: 50px; height: 50px;"" onerror=""this.src='http://192.168.0.227:92/MultiMedia/Sconosciuto.png'""  />"
				Dim sMeteo As String = " '" & MetteMaiuscoleDopoOgniSpazio("" & Rec("Tempo").Value) & "'<br />Gradi: " & Rec("Gradi").Value & " Umidità: " & Rec("Umidita").Value & " Pressione: " & Rec("Pressione").Value
				Dim Casa As String = "" & Rec("Casa").Value

				Dim Meteo As New StringBuilder

				Meteo.Append("<div style=""width: 100%;"">")
				Meteo.Append("<div style=""width: 15%; float: left; text-align: center;"">")
				Meteo.Append(immMeteo)
				Meteo.Append("</div>")
				Meteo.Append("<div style=""width: 70%; float: left; text-align: center;"">")
				Meteo.Append(sMeteo)
				Meteo.Append("</div>")
				Meteo.Append("<div style=""width: 15%; float: left; text-align: center;"">")
				Meteo.Append(immMeteo)
				Meteo.Append("</div>")
				Meteo.Append("</div>")

				Filone = Filone.Replace("***PARTITA***", "" & idPartita)
				Filone = Filone.Replace("***TIPOLOGIA***", "" & Rec("Tipologia").Value)
				Filone = Filone.Replace("***DATA ORA***", "" & Rec("DataOra").Value)
				If "" & Rec("Casa").Value = "E" Then
					Filone = Filone.Replace("***CAMPO***", "Campo esterno: " & Rec("CampoEsterno").Value)
					Filone = Filone.Replace("***INDIRIZZO***", Rec("CampoIndirizzo").Value)
				Else
					If (Rec("Casa").Value = "N") Then
						Filone = Filone.Replace("***CAMPO***", "" & Rec("CampoA").Value)
						Filone = Filone.Replace("***INDIRIZZO***", "" & Rec("CampoIndirizzo").Value)
					Else
						Filone = Filone.Replace("***CAMPO***", "" & "" & Rec("CampoSquadra").Value)
						Filone = Filone.Replace("***INDIRIZZO***", "" & Rec("Indirizzo").Value)
					End If
				End If
				Filone = Filone.Replace("***METEO***", "" & Meteo.ToString)

				Dim Notelle As String = "" & Rec("Note").Value

				If Notelle <> "" Then
					Notelle = Notelle.Replace("**BR**", "<br />")

					Filone = Filone.Replace("***NOTE***", "NOTE: <hr />" & Notelle)

					altezzaReport += 70
				Else
					Filone = Filone.Replace("***NOTE***", "")
				End If

				Dim CiSonoGiochetti As Boolean = False
				Dim Giochetti() As String = {}

				If Rec("ShootOut").Value = "S" Then
					If Rec("RisGiochetti").Value.ToString.Contains("%") And Rec("RisGiochetti").Value.ToString.Trim <> "%" Then
						Giochetti = Rec("RisGiochetti").Value.ToString.Split("%")
						Filone = Filone.Replace("***TIT RIS GIOCHETTI***", "Risultato Shoot Out:")
						Filone = Filone.Replace("***TRATTINO2***", "-")
						Filone = Filone.Replace("***RIS 1G***", Val(Giochetti(0)) + Val(Giochetti(2)))
						Filone = Filone.Replace("***RIS 2G***", Val(Giochetti(1)) + Val(Giochetti(3)))

						CiSonoGiochetti = True
					Else
						Filone = Filone.Replace("***TIT RIS GIOCHETTI***", "")
						Filone = Filone.Replace("***TRATTINO2***", "")
						Filone = Filone.Replace("***RIS 1G***", "")
						Filone = Filone.Replace("***RIS 2G***", "")
					End If
				Else
					Filone = Filone.Replace("***TIT RIS GIOCHETTI***", "")
					Filone = Filone.Replace("***TRATTINO2***", "")
					Filone = Filone.Replace("***RIS 1G***", "")
					Filone = Filone.Replace("***RIS 2G***", "")
				End If

				Dim NomeSquadra As String = ""
				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & Val(ss(1)).ToString
				Rec2 = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec2) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
				Else
					If Rec2.Eof() Then
					Else
						NomeSquadra = "" & Rec2("Descrizione").Value
					End If
				End If
				Rec2.Close()

				Dim ImmAll As String = PathBaseMultimedia & "/" & NomeSquadra & "/Allenatori/" & idAnno & "_" & Rec("idAllenatore").Value & ".kgb"
				ImmAll = DecriptaImmagine(MP, ImmAll)
				'Return ImmAll

				Filone = Filone.Replace("***IMMAGINE ALL***", ImmAll)
				Filone = Filone.Replace("***ALLENATORE***", "" & Rec("Allenatore").Value)

				Dim Imm1 As String = PathBaseMultimedia & "/" & NomeSquadra & "/Categorie/" & idAnno & "_" & Rec("idCategoria").Value & ".kgb"
				Imm1 = DecriptaImmagine(MP, Imm1)
				Dim Imm2 As String = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & Rec("idAvversario").Value & ".kgb"
				Imm2 = DecriptaImmagine(MP, Imm2)

				If Casa = "S" Then
					Filone = Filone.Replace("***IMMAGINE SQ1***", Imm1)
					Filone = Filone.Replace("***SQUADRA 1***", "" & Rec("Categoria").Value)

					Filone = Filone.Replace("***IMMAGINE SQ2***", Imm2)
					Filone = Filone.Replace("***SQUADRA 2***", "" & Rec("Avversario").Value)
				Else
					Filone = Filone.Replace("***IMMAGINE SQ2***", Imm2)
					Filone = Filone.Replace("***SQUADRA 2***", Rec("Avversario").Value)

					Filone = Filone.Replace("***IMMAGINE SQ1***", Imm1)
					Filone = Filone.Replace("***SQUADRA 1***", Rec("Categoria").Value)
				End If

				Dim Tempi As String = ""
				Select Case Rec("Tempi").Value
					Case 1
						Tempi = "Primo tempo: " & Rec("Tempo1Tempo").Value
					Case 2
						Tempi = "Primo tempo: " & Rec("Tempo1Tempo").Value & " Secondo tempo: " & Rec("Tempo2Tempo").Value
					Case 3
						Tempi = "Primo tempo: " & Rec("Tempo1Tempo").Value & " Secondo tempo: " & Rec("Tempo2Tempo").Value & " Terzo Tempo: " & Rec("Tempo3Tempo").Value
				End Select

				Filone = Filone.Replace("***TEMPI DI GIOCO***", Tempi)

				Dim RisultatoATempi As String = "" & Rec("RisultatoATempi").Value.ToString.Trim

				Rec.Close()

				' Arbitro
				Sql = "Select Arbitri.idArbitro, Arbitri.Cognome, Arbitri.Nome " &
					"FROM(Partite INNER JOIN ArbitriPartite On Partite.idPartita = ArbitriPartite.idPartita) " &
					"INNER Join Arbitri ON ArbitriPartite.idArbitro = Arbitri.idArbitro " &
					"Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura arbitro"
				Else
					If Not Rec.Eof() Then
						Dim PathArb As String = PathBaseImmagini & "/Arbitri/" & Rec("idArbitro").Value & ".kgb"
						PathArb = DecriptaImmagine(MP, PathArb)

						Filone = Filone.Replace("***IMMAGINE ARB***", PathArb)
						Filone = Filone.Replace("***ARBITRO***", Rec("Cognome").Value & " " & Rec("Nome").Value)
					Else
						Filone = Filone.Replace("***IMMAGINE ARB***", PathBaseImmScon)
						Filone = Filone.Replace("***ARBITRO***", "Arbitro non impostato")
					End If
				End If
				Filone = Filone.Replace("***IMMAGINE ARB***", "")
				Filone = Filone.Replace("***ARBITRO***", "")

				' Dirigenti
				Dim Dirigenti As New StringBuilder

				Sql = "SELECT Dirigenti.idDirigente, Dirigenti.Cognome, Dirigenti.Nome " &
						"FROM (Partite INNER JOIN DirigentiPartite ON Partite.idPartita = DirigentiPartite.idPartita) INNER JOIN Dirigenti ON DirigentiPartite.idDirigente = Dirigenti.idDirigente " &
						"Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita & " And Dirigenti.idAnno=" & idAnno
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura dirigenti"
				Else
					Dirigenti.Append("<table style=""width: 99%; text-align: center;"" cellpadding=""0"" cellspacing=""0"">")

					Do Until Rec.Eof()
						Dim Path As String = PathBaseImmagini & "/" & NomeSquadra & "/Dirigenti/" & idAnno & "_" & Rec("idDirigente").Value & ".kgb"
						Path = DecriptaImmagine(MP, Path)

						Dirigenti.Append("<tr>")
						Dirigenti.Append("<td>")
						Dirigenti.Append("<img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
						Dirigenti.Append("</td>")
						Dirigenti.Append("<td>")
						Dirigenti.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec("Cognome").Value & " " & Rec("Nome").Value & "</span>")
						Dirigenti.Append("</td>")
						Dirigenti.Append("</tr>")
						Dirigenti.Append(vbCrLf)

						Rec.MoveNext
					Loop

					Dirigenti.Append("</table>")
				End If

				Filone = Filone.Replace("***DIRIGENTE***", Dirigenti.ToString)

				Rec.Close()

				' Ammoniti / Espulsi
				Dim Ammoniti As New List(Of AmmoEspu)
				Dim Espulsi As New List(Of AmmoEspu)

				Sql = "Select idGiocatore, Descrizione, Minuto, idTempo From " &
					"EventiPartita " &
					"Left Join Eventi On EventiPartita.idEvento = Eventi.idEvento " &
					"Where idAnno = " & idAnno & " And idPartita = " & idPartita & " And (Upper(Descrizione)='AMMONITO' Or Upper(Descrizione)='ESPULSO')"
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura convocati"
				Else
					Do Until Rec.Eof()
						Dim a As New AmmoEspu
						a.idGiocatore = Rec("idGiocatore").Value
						a.Minuto = Rec("Minuto").Value
						a.Tempo = Rec("idTempo").Value

						If Rec("Descrizione").Value.ToString.ToUpper.Trim = "AMMONITO" Then
							Ammoniti.Add(a)
						Else
							Espulsi.Add(a)
						End If

						Rec.MoveNext
					Loop
					Rec.Close()
				End If

				' Convocati
				Sql = "SELECT Convocati.idGiocatore, Giocatori.NumeroMaglia, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione As Ruolo " &
					"FROM Partite " &
					"LEFT JOIN Convocati ON Partite.idPartita = Convocati.idPartita " &
					"LEFT JOIN Giocatori On Convocati.idGiocatore = Giocatori.idGiocatore AND Partite.idAnno = Giocatori.idAnno " &
					"LEFT JOIN [Generale].[dbo].[Ruoli] ON Giocatori.idRuolo = Ruoli.idRuolo " &
					"Where Partite.idAnno=" & idAnno & " And PArtite.idPartita=" & idPartita & " " &
					"Order By Ruoli.idRuolo"
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura convocati"
				Else
					Dim Convocati As New StringBuilder
					Dim Colore As String = "#aaa"

					altezzaConvocati += 40

					Convocati.Append("<table style=""width: 99%; text-align: center;"" cellpadding=""0"" cellspacing=""0"">")

					Do Until Rec.Eof()
						Dim C As String = Rec("Cognome").Value & " " & Rec("Nome").Value & "<br />" & Rec("Ruolo").Value
						Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
						Path = DecriptaImmagine(MP, Path)

						Dim messoTit As Boolean = False

						' Titolari
						Convocati.Append("<tr style=""background-color: " & Colore & """>")
						Convocati.Append("<td>")
						For Each inf As Integer In InFormazione
							If inf = Rec("idGiocatore").Value Then
								Convocati.Append("<img src=""" & PathBaseImmagini & "/Titolare.png" & """ width=""30px"" height=""30px"">")
								' altezzaConvocati += 63
								messoTit = True
								Exit For
							End If
						Next
						If Not messoTit Then Convocati.Append("&nbsp;")
						Convocati.Append("</td>")

						Convocati.Append("<td>")
						If idCapitano = Rec("idGiocatore").Value Then
							Convocati.Append("<span class=""testo nero"" style=""font-size: 15px; color: green;"">C</span>")
						Else
							Convocati.Append("")
						End If
						Convocati.Append("</td>")

						' Sostituzione
						Dim messoSost As Boolean = False

						Convocati.Append("<td style=""text-align: center;"">")
						For Each s As Sostituz In Sostituzioni
							If s.idEntrato = Rec("idGiocatore").Value Then
								Convocati.Append("<img src=""" & PathBaseImmagini & "/Entrato.png" & """ width=""30px"" height=""30px"">")
								Convocati.Append("<br /><span class=""testo nero"" style=""font-size: 13px;"">Min. " & s.Minuto & "</span><br /><span class=""testo nero"" style=""font-size: 13px;"">T. " & s.Tempo & "°</span>")
								altezzaConvocati += 63
								messoSost = True
								Exit For
							Else
								If s.idUscito = Rec("idGiocatore").Value Then
									Convocati.Append("<img src=""" & PathBaseImmagini & "/Uscito.png" & """ width=""30px"" height=""30px"">")
									Convocati.Append("<br /><span class=""testo nero"" style=""font-size: 13px;"">Min. " & s.Minuto & "</span><br /><span class=""testo nero"" style=""font-size: 13px;"">T. " & s.Tempo & "°</span>")
									altezzaConvocati += 63
									messoSost = True
									Exit For
								End If
							End If
						Next
						If Not messoSost Then
							altezzaConvocati += 50
						End If
						Convocati.Append("<td>")

						' Ammonizioni
						Convocati.Append("<td style=""text-align: center;"">")
						For Each s As AmmoEspu In Ammoniti
							If s.idGiocatore = Rec("idGiocatore").Value Then
								Convocati.Append("<img src=""" & PathBaseImmagini & "/Giallo.png" & """ width=""30px"" height=""30px"">")
								Convocati.Append("<br /><span class=""testo nero"" style=""font-size: 13px;"">Min. " & s.Minuto & "</span><br /><span class=""testo nero"" style=""font-size: 13px;"">T. " & s.Tempo & "°</span>")
								Exit For
							End If
						Next
						Convocati.Append("<td>")

						' Espulsioni
						Convocati.Append("<td style=""text-align: center;"">")
						For Each s As AmmoEspu In Espulsi
							If s.idGiocatore = Rec("idGiocatore").Value Then
								Convocati.Append("<img src=""" & PathBaseImmagini & "/Rosso.png" & """ width=""30px"" height=""30px"">")
								Convocati.Append("<br /><span class=""testo nero"" style=""font-size: 13px;"">Min. " & s.Minuto & "</span><br /><span class=""testo nero"" style=""font-size: 13px;"">T. " & s.Tempo & "°</span>")
								Exit For
							End If
						Next
						Convocati.Append("<td>")

						Convocati.Append("<td>")
						Convocati.Append("<img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
						Convocati.Append("</td>")
						Convocati.Append("<td style=""text-align: center;"">")
						Convocati.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec("NumeroMaglia").Value & "</span>")
						Convocati.Append("</td>")
						Convocati.Append("<td style=""text-align: left;"">")
						Convocati.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & C & "</span>")
						Convocati.Append("</td>")
						Convocati.Append("</tr>")
						Convocati.Append(vbCrLf)

						If Colore = "#aaa" Then Colore = "#fff" Else Colore = "#aaa"

						Rec.MoveNext
					Loop
					Rec.Close()

					Convocati.Append("</table>")

					Filone = Filone.Replace("***CONVOCATI***", Convocati.ToString)

					Dim QuantiGoal2 As Integer = 0

					Sql = "Select * From RisultatiAggiuntivi Where idPartita=" & idPartita
					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ok = False
						Ritorno = "Problemi lettura rislutati aggiuntivi: " & Sql
					Else
						Dim GoalAvv1Tempi As Integer = Val("" & Rec("GoalAvvPrimoTempo").Value)
						Dim GoalAvv2Tempi As Integer = Val("" & Rec("GoalAvvSecondoTempo").Value)
						Dim GoalAvv3Tempi As Integer = Val("" & Rec("GoalAvvTerzoTempo").Value)

						QuantiGoal2 = GoalAvv1Tempi + GoalAvv2Tempi + GoalAvv3Tempi
					End If
					Rec.Close()

					' Marcatori
					If TipoDB = "SQLSERVER" Then
						Sql = "Select * From (" &
							"SELECT RisultatiAggiuntiviMarcatori.Minuto, Giocatori.NumeroMaglia, Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione As Ruolo, RisultatiAggiuntiviMarcatori.idTempo, RisultatiAggiuntiviMarcatori.Rigore " &
							"FROM ((Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
							"INNER JOIN Giocatori ON (Partite.idAnno = Giocatori.idAnno) And (RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore)) " &
							"INNER JOIN [Generale].[dbo].[Ruoli] ON Giocatori.idRuolo = Ruoli.idRuolo " &
							"Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita & " " &
							"Union ALL " &
							"SELECT RisultatiAggiuntiviMarcatori.Minuto, '', -1, 'Autorete', '', '' As Ruolo, RisultatiAggiuntiviMarcatori.idTempo, RisultatiAggiuntiviMarcatori.Rigore " &
							"FROM Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita " &
							"Where Partite.idAnno = " & idAnno & " And Partite.idPartita = " & idPartita & " And IdGiocatore = -1 " &
							"Union All " &
							"Select value As Minuto, '', -2, 'Avversario', '', '' As Ruolo, 1 As idTempo, 'N' As Rigore From RisultatiAvversariMinuti " &
							"CROSS APPLY STRING_SPLIT(Minuti, ';') " &
							"Where idPartita = " & idPartita & " And idTempo = 1 And value <> '' " &
							"Union All " &
							"Select value As Minuto, '', -2, 'Avversario', '', '' As Ruolo, 2 As idTempo, 'N' As Rigore From RisultatiAvversariMinuti " &
							"CROSS APPLY STRING_SPLIT(Minuti, ';') " &
							"Where idPartita = " & idPartita & " And idTempo = 2 And value <> '' " &
							"Union All " &
							"Select value As Minuto, '', -2, 'Avversario', '', '' As Ruolo, 3 As idTempo, 'N' As Rigore From RisultatiAvversariMinuti " &
							"CROSS APPLY STRING_SPLIT(Minuti, ';') " &
							"Where idPartita = " & idPartita & " And idTempo = 3 And value <> '' " &
							") A " &
							"Order By idTempo, Minuto"
					Else
						Sql = "Delete From RisAvvMin"
						Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
						If Ritorno <> "OK" Then
							Return Ritorno
						End If

						Dim Rec22 As Object

						For i As Integer = 1 To 3
							Sql = "Select * From RisultatiAvversariMinuti Where idPartita = " & idPartita & " And idTempo = " & i
							Rec22 = Conn.LeggeQuery(MP, Sql, Connessione)
							If TypeOf (Rec22) Is String Then
								Ritorno = Rec22
								Return Ritorno
							Else
								If Not Rec22.Eof() Then
									Dim minuti() As String = Rec22("Minuti").Value.split(";")
									For Each m As String In minuti
										If m <> "" Then
											Sql = "Insert Into RisAvvMin Values (" & i & ", " & m & ")"
											Ritorno = Conn.EsegueSql(MP, Sql, Connessione)
											If Ritorno <> "OK" Then
												Return Ritorno
											End If
										End If
									Next
								End If
							End If
						Next

						Sql = "Select * From (" &
							"SELECT RisultatiAggiuntiviMarcatori.Minuto, Giocatori.NumeroMaglia, Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione As Ruolo, RisultatiAggiuntiviMarcatori.idTempo, RisultatiAggiuntiviMarcatori.Rigore " &
							"FROM ((Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
							"INNER JOIN Giocatori ON (Partite.idAnno = Giocatori.idAnno) And (RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore)) " &
							"INNER JOIN [Generale].[dbo].[Ruoli] ON Giocatori.idRuolo = Ruoli.idRuolo " &
							"Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita & " " &
							"Union ALL " &
							"SELECT RisultatiAggiuntiviMarcatori.Minuto, '', -1, 'Autorete', '', '' As Ruolo, RisultatiAggiuntiviMarcatori.idTempo, RisultatiAggiuntiviMarcatori.Rigore " &
							"FROM Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita " &
							"Where Partite.idAnno = " & idAnno & " And Partite.idPartita = " & idPartita & " And IdGiocatore = -1 " &
							"Union All " &
							"Select Minuto, '', -2, 'Avversario', '', '' As Ruolo, idTempo, 'N' As Rigore From RisAvvMin Where idTempo = 1 " &
							"Union All " &
							"Select Minuto, '', -2, 'Avversario', '', '' As Ruolo, idTempo, 'N' As Rigore From RisAvvMin Where idTempo = 2 " &
							"Union All " &
							"Select Minuto, '', -2, 'Avversario', '', '' As Ruolo, idTempo, 'N' As Rigore From RisAvvMin Where idTempo = 3 " &
							") A " &
							"Order By idTempo, Minuto"
					End If
					' Return Sql

					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ok = False
						Ritorno = "Problemi lettura marcatori: " & Sql
					Else
						Dim Marc() As String = {}
						Dim QuantiGoal As Integer = 0
						Dim QuantiGoal1 As Integer = 0
						'Dim QuantiGoal2 As Integer = 0

						Do Until Rec.Eof()
							ReDim Preserve Marc(QuantiGoal)
							Dim Minuto As String = "" & Rec("Minuto").Value
							If Minuto.Length = 1 Then Minuto = "0" & Minuto
							Marc(QuantiGoal) = "0" & Rec("idTempo").Value & ";" & Minuto & ";" & Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("Ruolo").Value & ";" & Rec("Rigore").Value & ";"
							If Rec("idGiocatore").Value <> -2 Then
								QuantiGoal1 += 1
							End If
							QuantiGoal += 1

							Rec.MoveNext
						Loop
						Rec.Close()

						For i As Integer = 0 To Marc.Length - 1
							For k As Integer = 0 To Marc.Length - 1
								If i <> k Then
									If Marc(i) < Marc(k) Then
										Dim Appo As String = Marc(i)
										Marc(i) = Marc(k)
										Marc(k) = Appo
									End If
								End If
							Next
						Next

						'' Tempi goal avversari
						'Sql = "Select * From RisultatiAvversariMinuti Where idPartita=" & idPartita & " Order by idTempo"
						'Rec = Conn.LeggeQuery(MP, Sql, Connessione)
						'If TypeOf (Rec) Is String Then
						'	Ok = False
						'	Ritorno = "Problemi lettura marcatori: " & Sql
						'Else
						'	Do Until Rec.Eof()
						'		Select Case Rec("idTempo").Value
						'			Case 1
						'				Dim ga1() As String = Rec("Minuti").Value.Split(";")

						'				For Each g As String In ga1
						'					If g <> "" Then
						'						ReDim Preserve Marc(QuantiGoal)
						'						Marc(QuantiGoal) = "01;" & Format(Val(g), "00") & ";;Goal avversario;;;"
						'						QuantiGoal2 += 1
						'						QuantiGoal += 1
						'					End If
						'				Next
						'			Case 2
						'				Dim ga2() As String = Rec("Minuti").Value.Split(";")

						'				For Each g As String In ga2
						'					If g <> "" Then
						'						ReDim Preserve Marc(QuantiGoal)
						'						Marc(QuantiGoal) = "02;" & Format(Val(g), "00") & ";;Goal avversario;;;"
						'						QuantiGoal2 += 1
						'						QuantiGoal += 1
						'					End If
						'				Next
						'			Case 3
						'				Dim ga3() As String = Rec("Minuti").Value.Split(";")

						'				For Each g As String In ga3
						'					If g <> "" Then
						'						ReDim Preserve Marc(QuantiGoal)
						'						Marc(QuantiGoal) = "03;" & Format(Val(g), "00") & ";;Goal avversario;;;"
						'						QuantiGoal2 += 1
						'						QuantiGoal += 1
						'					End If
						'				Next
						'		End Select

						'		Rec.MoveNext
						'	Loop
						'	Rec.Close()
						'End If

						Dim GoalPropri As Integer = 0
						Dim GoalAvversari As Integer = 0
						Dim NomiCampi() As String = {"", "GoalAvvPrimoTempo", "GoalAvvSecondoTempo", "GoalAvvTerzoTempo"}
						Dim RisProprio As Integer = 0
						Dim RisAvversario As Integer = 0

						If RisultatoATempi = "S" Then
							For i As Integer = 1 To 3
								Sql = "Select " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " From RisultatiAggiuntiviMarcatori Where idPartita=" & idPartita & " And idTempo=" & i
								Rec = Conn.LeggeQuery(MP, Sql, Connessione)
								'If Rec(0).Value Is DBNull.Value Then
								'	GoalPropri = 0
								'Else
								GoalPropri = Rec(0).Value
								'End If
								Rec.Close()
								Sql = "Select " & IIf(TipoDB = "SQLSERVER", "Isnull(Sum(" & NomiCampi(i) & "),0)", "Colaesce(Sum(" & NomiCampi(i) & "),0)") & " From RisultatiAggiuntivi Where idPartita=" & idPartita & " And " & NomiCampi(i) & "<>-1"
								Rec = Conn.LeggeQuery(MP, Sql, Connessione)
								'If Rec(0).Value Is DBNull.Value Then
								'	GoalAvversari = 0
								'Else
								GoalAvversari = Rec(0).Value
								'End If
								Rec.Close()

								If GoalPropri > GoalAvversari Then
									RisProprio += 1
								Else
									If GoalPropri < GoalAvversari Then
										RisAvversario += 1
									Else
										RisProprio += 1
										RisAvversario += 1
									End If
								End If
							Next
						Else
							Sql = "Select * From RisultatiAggiuntivi Where idPartita=" & idPartita
							Rec = Conn.LeggeQuery(MP, Sql, Connessione)
							If Rec.Eof() = False Then
								GoalAvversari = Val("" & Rec(NomiCampi(1)).Value) + Val("" & Rec(NomiCampi(2)).Value) + Val("" & Rec(NomiCampi(3)).Value)
							End If
							Rec.Close()
						End If

						If CiSonoGiochetti Then
							If Val(Giochetti(0)) > Val(Giochetti(1)) Then
								RisProprio += 1
							Else
								If Val(Giochetti(0)) < Val(Giochetti(1)) Then
									RisAvversario += 1
								Else
									RisProprio += 1
									RisAvversario += 1
								End If
							End If
						End If

						Dim CiSonoRigori As Boolean = False
						Dim Rigoristi As New List(Of String)
						Dim RigoriAvv As String = ""
						Dim RigoriSegnatiPropri As Integer = 0
						Dim RigoriSegnatiAvversari As Integer = 0
						Dim RigoriSbagliatiAvversari As Integer = 0

						Sql = "SELECT RigoriPropri.idGiocatore, RigoriPropri.idRigore, Ruoli.Descrizione, " & IIf(TipoDB = "SQLSERVER", "Giocatori.Cognome + ' ' + Giocatori.Nome", "Concat(Giocatori.Cognome, ' ', Giocatori.Nome)") & " As Giocatore, " &
							"Giocatori.NumeroMaglia, RigoriPropri.Termine From ((RigoriPropri " &
							"Left Join Giocatori On RigoriPropri.idGiocatore=Giocatori.idGiocatore And RigoriPropri.idAnno = Giocatori.idAnno) " &
							"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo = Ruoli.idRuolo) " &
							"Where RigoriPropri.idAnno=" & idAnno & " And idPartita=" & idPartita & " " &
							"Order By RigoriPropri.idRigore"
						Rec2 = Conn.LeggeQuery(MP, Sql, Connessione)
						If TypeOf (Rec2) Is String Then
							Ok = False
							Ritorno = "Problemi lettura rigori"
						Else
							If Not Rec2.Eof() Then
								CiSonoRigori = True

								Do Until Rec2.Eof()
									Dim Termine As String = ""
									Dim Colore2 As String = ""

									If Rec2("Termine").Value = "1" Then
										Termine = "Segnato"
										Colore2 = "verde"
										RigoriSegnatiPropri += 1
									Else
										If Rec2("Termine").Value = "0" Then
											Termine = "Sbagliato"
											Colore2 = "rosso"
										End If
									End If
									If Termine <> "" Then
										' Rigoristi.Add("<span class=""testo " & Colore & """ style=""font-size: 15px;"">Rigore " & Rec2("idRigore").Value & ": " & Rec2("Giocatore").Value & " (" & Rec2("Descrizione").Value & ") - " & Termine & "</span>")
										Rigoristi.Add(Colore2 & ";" & Rec2("idRigore").Value & ";;" & Rec2("Giocatore").Value & ";" & Rec2("Descrizione").Value & "; " & Termine & ";" & Rec2("idGiocatore").Value & ";")
									End If

									Rec2.MoveNext
								Loop
								Rec2.Close()
							End If
						End If

						'If CiSonoRigori Then
						'	Sql = "Select * From RigoriAvversari Where idAnno=" & idAnno & " And idPartita=" & idPartita
						'	Rec2 = Conn.LeggeQuery(MP, Sql, Connessione)
						'	If TypeOf (Rec2) Is String Then
						'	Else
						'		If Not Rec2.Eof() Then
						'			RigoriSegnatiAvversari += Val(Rec2("Segnati").Value)
						'			RigoriSbagliatiAvversari += Val(Rec2("Sbagliati").Value)

						'			RigoriAvv = Rec2("Segnati").Value & "!" & Rec2("Sbagliati").Value & "!"
						'		End If
						'	End If

						'	Dim Rigori As String = "<span class=""testo blu"" style=""font-size: 15px;"">RISULTATO DOPO I TEMPI REGOLAMENTARI: " & QuantiGoal1 & " - " & QuantiGoal2 & "</span><br /><br />"

						'	Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">RIGORI PROPRI:</span><hr />"

						'	Rigori &= "<table style=""width: 99%; text-align: center;"">"
						'	For Each s As String In Rigoristi
						'		Dim c() As String = s.Split(";")
						'		Dim Path2 As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & c(6) & ".kgb"
						'		Path2 = DecriptaImmagine(Path2)

						'		Rigori &= "<tr>"
						'		Rigori &= "<td align=""left"">"
						'		Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">Rigore " & c(1) & "</span>"
						'		Rigori &= "</td>"
						'		Rigori &= "<td>"
						'		Rigori &= "<img src=""" & Path2 & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />"
						'		Rigori &= "</td>"
						'		Rigori &= "<td align=""center"">"
						'		Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">" & c(3) & "</span>"
						'		Rigori &= "</td>"
						'		Rigori &= "<td align=""center"">"
						'		Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">" & c(4) & "</span>"
						'		Rigori &= "</td>"
						'		Rigori &= "<td align=""center"">"
						'		Rigori &= "<span class=""testo " & c(0) & """ style=""font-size: 15px;"">" & c(5) & "</span>"
						'		Rigori &= "</td>"
						'		Rigori &= "</tr>"
						'	Next
						'	Rigori &= "</table>"

						'	Rigori &= "<br /><span class=""testo blu"" style=""font-size: 15px;"">RIGORI AVVERSARI:</span><hr />"
						'	Rigori &= "<span class=""testo rosso"" style=""font-size: 15px;"">Segnati: " & RigoriSegnatiAvversari & "</span><br />"
						'	Rigori &= "<span class=""testo verde"" style=""font-size: 15px;"">Sbagliati: " & RigoriSbagliatiAvversari & "</span><hr />"

						'	Filone = Filone.Replace("***RIGORI***", Rigori)
						Filone = Filone.Replace("***RIGORI***", "")

						'	If RisultatoATempi = "S" Then
						'		RisProprio += RigoriSegnatiPropri
						'		RisAvversario += RigoriSegnatiAvversari
						'	Else
						'		If (RigoriSegnatiPropri > RigoriSegnatiAvversari) Then
						'			RisProprio += 1
						'		Else
						'			If (RigoriSegnatiPropri < RigoriSegnatiAvversari) Then
						'				RisAvversario += 1
						'			End If
						'		End If
						'	End If
						'Else
						'	Filone = Filone.Replace("***RIGORI***", "")
						'End If

						If RisultatoATempi = "S" Then
							Filone = Filone.Replace("***TIT RIS TEMPI***", "Risultato a tempi:")
							Filone = Filone.Replace("***RIS 1T***", RisProprio.ToString.Trim)
							Filone = Filone.Replace("***RIS 2T***", RisAvversario.ToString.Trim)
							Filone = Filone.Replace("***TRATTINO1***", "-")
						Else
							Filone = Filone.Replace("***RIS 1T***", "")
							Filone = Filone.Replace("***RIS 2T***", "")
							Filone = Filone.Replace("***TIT RIS TEMPI***", "")
							Filone = Filone.Replace("***TRATTINO1***", "")
						End If

						Dim Marcatori As New StringBuilder

						Marcatori.Append("<table style=""width: 99%; text-align: center;"" cellpadding=""0"" cellspacing=""0"">")
						Marcatori.Append(vbCrLf)
						Marcatori.Append("<tr>")
						Marcatori.Append("<td>")
						Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Rig.</span>")
						Marcatori.Append("</td>")
						Marcatori.Append("<td>")
						Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">T.</span>")
						Marcatori.Append("</td>")
						Marcatori.Append("<td>")
						Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Min.</span>")
						Marcatori.Append("</td>")
						Marcatori.Append("<td>")
						Marcatori.Append("</td>")
						Marcatori.Append("<td>")
						Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Giocatore</span>")
						Marcatori.Append("</td>")
						'Marcatori.Append("<td>")
						'Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Ruolo</span>")
						'Marcatori.Append("</td>")
						Marcatori.Append("</tr>")
						Marcatori.Append(vbCrLf)

						Dim OldTempo As String = ""

						Colore = "#aaa"
						For Each m As String In Marc
							Dim Mm() As String = m.Split(";")

							If OldTempo <> Mm(0) Then
								Marcatori.Append("<tr>")
								Marcatori.Append("<td>")
								Marcatori.Append("<hr />")
								Marcatori.Append("</td>")
								Marcatori.Append("<td>")
								Marcatori.Append("<hr />")
								Marcatori.Append("</td>")
								Marcatori.Append("<td>")
								Marcatori.Append("<hr />")
								Marcatori.Append("</td>")
								Marcatori.Append("<td>")
								Marcatori.Append("<hr />")
								Marcatori.Append("</td>")
								Marcatori.Append("<td>")
								Marcatori.Append("<hr />")
								Marcatori.Append("</td>")
								'Marcatori.Append("<td>")
								'Marcatori.Append("<hr />")
								'Marcatori.Append("</td>")
								Marcatori.Append("</tr>")
								Marcatori.Append(vbCrLf)

								OldTempo = Mm(0)
							End If

							Dim Path As String

							If m.Contains("Goal avversario") Then
								Path = PathBaseMultimedia & "/goal.png"
							Else
								If m.Contains("Autorete") Then
									Path = PathBaseMultimedia & "/autorete.png"
								Else
									If m.Contains("Avversario") Then
										If Casa = "S" Then
											Path = Imm1
										Else
											Path = Imm2
										End If
									Else
										Path = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Mm(2) & ".kgb"
										Path = DecriptaImmagine(MP, Path)
									End If
								End If
							End If

							Marcatori.Append("<tr style=""background-color: " & Colore & """>")
							Marcatori.Append("<td>")
							If Mm(6) = "S" Then
								Marcatori.Append("<span class=""testo nero"" style=""font-size: 15px; font-weight: bold; color: red;"">*</span>")
							Else
								Marcatori.Append("")
							End If
							Marcatori.Append("</td>")
							Marcatori.Append("<td>")
							Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(0) & "</span>")
							Marcatori.Append("</td>")
							Marcatori.Append("<td>")
							Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(1) & "°</span>")
							Marcatori.Append("</td>")
							Marcatori.Append("<td>")
							Marcatori.Append("<img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
							Marcatori.Append("</td>")
							Marcatori.Append("<td style=""text-align: left;"">")
							Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(3) & " " & Mm(4) & "<br />" & Mm(5) & "</span>")
							Marcatori.Append("</td>")
							'Marcatori.Append("<td>")
							'Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(5) & "</span>")
							'Marcatori.Append("</td>")
							Marcatori.Append("</tr>")
							Marcatori.Append(vbCrLf)

							If Colore = "#aaa" Then Colore = "#fff" Else Colore = "#aaa"
						Next

						Marcatori.Append("</table>")

						Filone = Filone.Replace("***MARCATORI***", Marcatori.ToString)
						'Filone = Marcatori.ToString

						' Eventi
						Dim Eventi As New StringBuilder
						Dim Evento As New List(Of String)
						Dim QuantiEventi As New List(Of Integer)

						Eventi.Append("<table style=""width: 99%; text-align: center;"" cellpadding=""0"" cellspacing=""0"">")
						Eventi.Append(vbCrLf)

						If TipoDB = "SQLSERVER" Then
							Sql = "SELECT EventiPartita.idTempo, EventiPartita.Minuto, Eventi.Descrizione, " &
								"IIf(IsNull(Giocatori.Cognome + ' ' + Giocatori.Nome, '') = '', 'Avversario', IsNull(Giocatori.Cognome + ' ' + Giocatori.Nome, '')" & " As Giocatore, " &
								"Giocatori.idGiocatore " &
								"FROM EventiPartita " &
								"LEFT JOIN Giocatori ON EventiPartita.idGiocatore = Giocatori.idGiocatore And EventiPartita.idAnno = Giocatori.idAnno " &
								"LEFT JOIN Eventi ON EventiPartita.idEvento = Eventi.idEvento " &
								"WHERE EventiPartita.idPartita=" & idPartita & " And EventiPartita.idAnno=" & idAnno
						Else
							Sql = "SELECT EventiPartita.idTempo, EventiPartita.Minuto, Eventi.Descrizione, " &
								"If(Trim(Concat(Coalesce(Giocatori.Cognome, ''), ' ', Coalesce(Giocatori.Nome,''))) = '', 'Avversario', Concat(Giocatori.Cognome, ' ', Giocatori.Nome))" & " As Giocatore, " &
								"Giocatori.idGiocatore " &
								"FROM EventiPartita " &
								"LEFT JOIN Giocatori ON EventiPartita.idGiocatore = Giocatori.idGiocatore And EventiPartita.idAnno = Giocatori.idAnno " &
								"LEFT JOIN Eventi ON EventiPartita.idEvento = Eventi.idEvento " &
								"WHERE EventiPartita.idPartita=" & idPartita & " And EventiPartita.idAnno=" & idAnno
						End If
						Rec2 = Conn.LeggeQuery(MP, Sql, Connessione)
						If TypeOf (Rec2) Is String Then
							Return Rec2
						Else
							Dim tempoAtt As String = ""
							Colore = "#aaa"

							Do Until Rec2.Eof = True
								Dim Path As String

								If Rec2("Giocatore").Value.Contains("Avversario") Then
									Path = PathBaseImmScon
								Else
									Path = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec2("idGiocatore").Value & ".kgb"
									Path = DecriptaImmagine(MP, Path)
								End If

								If Val(tempoAtt) <> Val(Rec2("idTempo").Value) Then
									If tempoAtt <> "" Then
										Eventi.Append("<tr><td colspan=""5""><hr /><span class=""testo rosso"" style=""font-size: 15px;"">Tempo " & Rec2("idTempo").Value & "</span><hr /></tr>")

										altezzaEventi += 55
									Else
										Eventi.Append("<tr><td colspan=""5""><span class=""testo rosso"" style=""font-size: 15px;"">Tempo " & Rec2("idTempo").Value & "</span><hr /></tr>")

										Eventi.Append("<tr>")
										Eventi.Append("<td>")
										Eventi.Append("<span class=""testo verde"" style=""font-size: 13px;"">Minuto</span>")
										Eventi.Append("</td>")
										Eventi.Append("<td>")
										Eventi.Append("<span class=""testo verde"" style=""font-size: 13px;"">Evento</span>")
										Eventi.Append("</td>")
										Eventi.Append("<td>")
										Eventi.Append("</td>")
										Eventi.Append("<td>")
										Eventi.Append("<span class=""testo verde"" style=""font-size: 13px;"">Giocatore</span>")
										Eventi.Append("</td>")
										Eventi.Append("</tr>")
										Eventi.Append(vbCrLf)

										altezzaEventi += 55
									End If

									tempoAtt = Val(Rec2("idTempo").Value)
								End If
								Eventi.Append("<tr style=""background-color: " & Colore & """>")
								'Eventi.Append("<td align=""right"">")
								'Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec2("idTempo").Value & "</span>")
								'Eventi.Append("</td>")
								Eventi.Append("<td align=""center"">")
								Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec2("Minuto").Value & "°</span>")
								Eventi.Append("</td>")
								Eventi.Append("<td align=""left"">")
								Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec2("Descrizione").Value & "</span>")
								Eventi.Append("</td>")
								Eventi.Append("<td>")
								Eventi.Append("<img src=""" & Path & """ style=""width: 30px; height: 30px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
								Eventi.Append("</td>")
								Eventi.Append("<td align=""left"">")

								Dim Nome As String = Rec2("Giocatore").Value
								Dim Cognome As String = ""

								If Nome.Contains(" ") Then
									Dim a As Integer = Nome.IndexOf(" ")

									Cognome = Mid(Nome, a + 1, Nome.Length).Trim
									Nome = Mid(Nome, 1, a).Trim
									Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Nome & "<br />" & Cognome & "</span>")
								Else
									Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Nome & "</span>")
								End If

								Eventi.Append("</td>")
								Eventi.Append("</tr>")
								Eventi.Append(vbCrLf)

								If Colore = "#aaa" Then Colore = "#fff" Else Colore = "#aaa"

								altezzaEventi += 29

								Dim OkEv As Boolean = False
								Dim q2 As Integer = 0

								For Each e As String In Evento
									If e = Rec2("Descrizione").Value Then
										OkEv = True
										QuantiEventi.Item(q2) += 1
										Exit For
									End If

									q2 += 1
								Next
								If Not OkEv Then
									Evento.Add(Rec2("Descrizione").Value)
									QuantiEventi.Add(1)
								End If

								Rec2.MoveNext
							Loop
							Rec2.Close()
						End If

						Eventi.Append("</table>")
						Filone = Filone.Replace("***RACCONTO***", Eventi.ToString)

						For i As Integer = 0 To QuantiEventi.Count - 1
							For k = i + 1 To QuantiEventi.Count - 1
								If QuantiEventi.Item(i) < QuantiEventi.Item(k) Then
									Dim Appo As Integer = QuantiEventi.Item(i)
									QuantiEventi.Item(i) = QuantiEventi.Item(k)
									QuantiEventi(k) = Appo

									Dim sAppo As String = Evento.Item(i)
									Evento.Item(i) = Evento.Item(k)
									Evento(k) = sAppo
								End If
							Next
						Next

						Dim q As Integer = 0
						Dim Riepilogo As New StringBuilder

						Riepilogo.Append("<div style=""width: 98%; border: 1px solid #999; border-radius: 3px; background-color: #fffccb; height: auto; padding: 3px; margin: 3px;"">")
						Riepilogo.Append("<table style=""width: 99%; text-align: center;"" cellpadding=""0"" cellspacing=""0"">")
						Riepilogo.Append(vbCrLf)
						Riepilogo.Append("<tr>")
						Riepilogo.Append("<th><span class=""testo verde"" style=""font-size: 13px;"">Evento</span>")
						Riepilogo.Append("</th>")
						Riepilogo.Append("<th><span class=""testo verde"" style=""font-size: 13px;"">Quanti</span>")
						Riepilogo.Append("</th>")
						Riepilogo.Append("</tr>")
						Riepilogo.Append(vbCrLf)

						For Each e As String In Evento
							Riepilogo.Append("<tr>")
							Riepilogo.Append("<td align=""left"">")
							Riepilogo.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & e & "</span>")
							Riepilogo.Append("</td>")
							Riepilogo.Append("<td align=""right"">")
							Riepilogo.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & QuantiEventi.Item(q) & "</span>")
							Riepilogo.Append("</td>")
							Riepilogo.Append("</tr>")
							Riepilogo.Append(vbCrLf)

							altezzaEventi += 29

							q += 1
						Next
						Riepilogo.Append("</table>")
						Riepilogo.Append("</div>")
						Riepilogo.Append(vbCrLf)

						Filone = Filone.Replace("***RIEPILOGO RACCONTO***", Riepilogo.ToString)

						If altezzaConvocati > altezzaEventi Then
							altezzaReport += altezzaConvocati
						Else
							altezzaReport += altezzaEventi
						End If

						' Risultato
						If CiSonoRigori Then
							QuantiGoal1 += RigoriSegnatiPropri
							QuantiGoal2 += RigoriSegnatiAvversari
						End If

						'If Casa = "S" Then
						Filone = Filone.Replace("***RIS 1***", QuantiGoal1)
						Filone = Filone.Replace("***RIS 2***", QuantiGoal2)
						'Else
						'    Filone = Filone.Replace("***RIS 1***", QuantiGoal2)
						'    Filone = Filone.Replace("***RIS 2***", QuantiGoal1)
						'End If

						If QuantiGoal1 > QuantiGoal2 Then
							Filone = Filone.Replace("***COLORE RIS***", "verde")
						Else
							If QuantiGoal1 < QuantiGoal2 Then
								Filone = Filone.Replace("***COLORE RIS***", "rosso")
							Else
								Filone = Filone.Replace("***COLORE RIS***", "nero")
							End If
						End If

						'Dim wb As New WebBrowser
						'wb.DocumentText = Filone
						'Dim doc2 As HtmlDocument = wb.Document
						'Dim width As Integer = wb.Document.Body.ScrollRectangle.Width
						'Dim height As Integer = wb.Document.Body.ScrollRectangle.Height

						gf.CreaAggiornaFile(NomeFileFinale, Filone)

						If TipoPDF = "WINDOWS" Then
							Dim pp As New pdfGest
							Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, PathLog, True,, altezzaReport)

							If Ritorno = "OK" Then
								Ritorno = NomeFileFinalePDF.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
							End If
						Else
							If Ritorno = "OK" Then
								' sudo html2pdf /var/www/html/inCalcio/Multimedia/allegati/0001_00012/Partite/Anno1/00155/155.html
								' /var/www/html/inCalcio/Multimedia/allegati/0001_00012/Partite/Anno1/00155/155.pdf
								Dim nomeFileDaConvertire As String = NomeFileFinale.Replace("\", "/")
								Dim nomeFileConvertito As String = nomeFileDaConvertire.Replace(".html", ".pdf")

								Ritorno = ConvertePDF(nomeFileDaConvertire, nomeFileConvertito)
								If Ritorno = "*" Then
									Ritorno = NomeFileFinalePDF.Replace(pathAllegati, urlAllegati).Replace("Multimedia", "Allegati").Replace("\", "/")
								End If

								' Ritorno = "html2pdf " & nomeFileDaConvertire & "  " & nomeFileConvertito
							End If
						End If
					End If
				End If
			Else
				Ok = False
				Ritorno = "ERROR: Nessun dato rilevato"
			End If
		End If

		Return Ritorno
	End Function

	Public Function ConvertePDF(NomeFileFinaleP As String, NomeFileFinalePDFP As String) As String
		Dim Ritorno As String = "*"
		Dim NomeFileFinale As String = NomeFileFinaleP
		Dim NomeFileFinalePDF As String = NomeFileFinalePDFP

		If TipoPATH <> "SQLSERVER" Then
			NomeFileFinale = NomeFileFinale.Replace("\", "/")
			NomeFileFinale = NomeFileFinale.Replace("//", "/")
			NomeFileFinale = NomeFileFinale.Replace("/\", "/")

			NomeFileFinalePDF = NomeFileFinalePDF.Replace("\", "/")
			NomeFileFinalePDF = NomeFileFinalePDF.Replace("//", "/")
			NomeFileFinalePDF = NomeFileFinalePDF.Replace("/\", "/")
		End If

		Try
			Dim processoConversione As Process = New Process()
			Dim pi As ProcessStartInfo = New ProcessStartInfo()
			pi.FileName = "html2pdf"
			pi.Arguments = NomeFileFinale & "  " & NomeFileFinalePDF
			pi.WindowStyle = ProcessWindowStyle.Normal
			processoConversione.StartInfo = pi
			processoConversione.StartInfo.UseShellExecute = False
			processoConversione.StartInfo.RedirectStandardOutput = True
			processoConversione.StartInfo.RedirectStandardError = True
			processoConversione.Start()

			Dim OutPutP As String = processoConversione.StandardOutput.ReadToEnd()

			processoConversione.WaitForExit()
		Catch ex As Exception
			Ritorno = StringaErrore & ": " & ex.Message
		End Try

		Return Ritorno
	End Function

	Public Function CriptaStringa(Stringa As String) As String
		'Dim Ancora As Boolean = True
		Dim Ritorno As String = ""

		'Do While Ancora = True
		'	Dim rnd1 As New Random(CInt(Date.Now.Ticks And &HFFFF))
		'	Dim Chiave As Integer = rnd1.Next(32) + 32
		'	Ritorno = Chr(Chiave)
		'	For i As Integer = 1 To 13
		'		Dim x As Integer = 64 + rnd1.Next(48)
		'		Ritorno &= Chr(x)
		'	Next
		'	Ritorno &= Chr(32 + Stringa.Length)
		'	Dim Contatore As Integer = 0
		'	For i As Integer = 1 To Stringa.Length
		'		Dim carattere As String = Mid(Stringa, i, 1)
		'		Dim ascii As Integer = Asc(carattere) + Chiave + Contatore
		'		Dim ascii2 As String = Chr(ascii)
		'		Ritorno &= ascii2
		'		Contatore += 2
		'	Next
		'	For i As Integer = 1 To 7
		'		Dim x As Integer = 64 + rnd1.Next(48)
		'		Ritorno &= Chr(x)
		'	Next

		'	If Not Ritorno.Contains(";") And Not Ritorno.Contains("'") Then
		'		Ancora = False
		'	End If
		'Loop

		Dim wrapper As New CryptEncrypt(CryptPasswordString)
		Ritorno = wrapper.EncryptData(Stringa)

		Return Ritorno
	End Function

	Public Function DecriptaStringa(Stringa As String) As String
		Dim Ritorno As String = ""
		'Dim Chiave As Integer = Asc(Mid(Stringa, 1, 1))
		'' Dim Chiave As Integer = Asc(car)
		'Dim Lunghezza As Integer = Asc(Mid(Stringa, 15, 1)) - 32
		'Dim Contatore As Integer = 0
		'For i As Integer = 16 To 16 + Lunghezza - 1
		'	Dim Car As String = Mid(Stringa, i, 1)
		'	Dim Car1 As Integer = Asc(Car) - Chiave - Contatore
		'	Dim c As String = Chr(Car1)
		'	Ritorno &= c
		'	Contatore += 2
		'Next

		Dim wrapper As New CryptEncrypt(CryptPasswordString)

		Try
			Ritorno = wrapper.DecryptData(Stringa)
		Catch ex As System.Security.Cryptography.CryptographicException
			Ritorno = ""
		End Try

		Return Ritorno
	End Function

	Public Function EliminaPartita(MP As String, Squadra As String, idAnno As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(MP, Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim sql As String
				Dim Ok As Boolean = True

				sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(MP, sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					sql = "delete from Partite Where idAnno = " & idAnno & " And idPartita = " & idPartita
					Ritorno = Conn.EsegueSql(MP, sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						sql = "delete from ArbitriPartite Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from Convocati Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from CoordinatePartite Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from EventiPartita Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from Marcatori Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from MeteoPartite Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RigoriAvversari Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RigoriPropri Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from Risultati Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RisultatiAggiuntivi Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RisultatiAggiuntiviMarcatori Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from TempiGoalAvversari Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from EventiConvocazioni Where idPartita = " & idPartita
						Ritorno = Conn.EsegueSql(MP, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If
				Else
					Ok = False
				End If

				If Ok Then
					sql = "commit"
					Dim Ritorno2 As String = Conn.EsegueSql(MP, sql, Connessione)
				Else
					sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(MP, sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function RitornaMeteo(Lat As String, Lon As String) As String
		Dim Ritorno As String = ""
		Dim Cosa As String = ""

		'If Lat = "undefined" Or Lon = "undefined" Or IsNumeric(Lat.Replace(",", ".")) = False Or IsNumeric(Lon.Replace(",", ".")) = False Then
		Cosa = "q=Rome,IT"
		'Else
		'	Cosa = "lat=" & Lat & "&lon=" & Lon
		'End If

		Try
			Dim url As String = "http://api.openweathermap.org/data/2.5/weather?" & Cosa & "&mode=xml&units=metric&lang=it&appid=1856b7a9244abb668591169ef0a34308"
			Dim request As WebRequest = WebRequest.Create(url)
			Dim response As WebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
			Dim reader As New StreamReader(response.GetResponseStream(), Encoding.UTF8)
			Dim dsResult As New DataSet()

			dsResult.ReadXml(reader)

			'Temperatura: dsResult.Tables(3).Rows(0)
			'Umidita: dsResult.Tables(5).Rows(0)
			'Pressione: dsResult.Tables(6).Rows(0)
			'Tempo:  dsResult.Tables(13).Rows(0)(1)

			Ritorno &= dsResult.Tables("weather").Rows(0)(1).ToString() & ";"

			'txtMinima.Text = dsResult.Tables(4).Rows(0)(1).ToString()
			'txtMassima.Text = dsResult.Tables(4).Rows(0)(2).ToString()
			Ritorno &= dsResult.Tables("temperature").Rows(0)(0).ToString() & ";"
			'txtSorge.Text = DateTime.Parse(dsResult.Tables(3).Rows(0)(0).ToString()).ToString("dd/MM/yyyy hh:mm:ss")
			'txtTramonta.Text = DateTime.Parse(dsResult.Tables(3).Rows(0)(1).ToString()).ToString("dd/MM/yyyy HH:mm:ss")
			Ritorno &= dsResult.Tables("humidity").Rows(0)(0).ToString() & ";"
			Ritorno &= dsResult.Tables("pressure").Rows(0)(0).ToString() & ";"
			'txtventoVelocita.Text = dsResult.Tables(8).Rows(0)(0).ToString() + " " + dsResult.Tables(8).Rows(0)(1).ToString()
			'txtDirezioneVento.Text = dsResult.Tables(9).Rows(0)(1).ToString() + "     " + dsResult.Tables(9).Rows(0)(2).ToString()
			'txtPrecipitazione.Text = dsResult.Tables(11).Rows(0)(0).ToString()

			Ritorno &= "http://openweathermap.org/img/w/" + dsResult.Tables("temperature").Rows(0)(2).ToString() + ".png" & ";"
		Catch ex As Exception
			Ritorno = StringaErrore & " " & ex.Message
		End Try

		Return Ritorno
	End Function

	Public Function convertNumberToReadableString(ByVal num As Long) As String
		Dim result As String = ""
		Dim [mod] As Long = 0
		Dim i As Single = 0
		Dim unita As String() = {"zero", "uno", "due", "tre", "quattro", "cinque", "sei", "sette", "otto", "nove", "dieci", "undici", "dodici", "tredici", "quattordici", "quindici", "sedici", "diciassette", "diciotto", "diciannove"}
		Dim decine As String() = {"", "dieci", "venti", "trenta", "quaranta", "cinquanta", "sessanta", "settanta", "ottanta", "novanta"}

		If num > 0 AndAlso num < 20 Then
			result = unita(num)
		Else

			If num < 100 Then
				[mod] = num Mod 10
				i = Int(num / 10)

				Select Case [mod]
					Case 0
						result = decine(i)
					Case 1
						result = decine(i).Substring(0, decine(i).Length - 1) & unita([mod])
					Case 8
						result = decine(i).Substring(0, decine(i).Length - 1) & unita([mod])
					Case Else
						result = decine(i) & unita([mod])
				End Select
			Else

				If num < 1000 Then
					[mod] = num Mod 100
					i = Int((num - [mod]) / 100)

					Select Case i
						Case 1
							result = "cento"
						Case Else
							result = unita(i) & "cento"
					End Select

					result = result & convertNumberToReadableString([mod])
				Else

					If num < 10000 Then
						[mod] = num Mod 1000
						i = Int((num - [mod]) / 1000)

						Select Case i
							Case 1
								result = "mille"
							Case Else
								result = unita(i) & "mila"
						End Select

						result = result & convertNumberToReadableString([mod])
					Else

						If num < 1000000 Then
							[mod] = num Mod 1000
							i = Int((num - [mod]) / 1000)

							Select Case (num - [mod]) / 1000
								Case Else

									If i < 20 Then
										result = unita(i) & "mila"
									Else
										result = convertNumberToReadableString(i) & "mila"
									End If
							End Select

							result = result & convertNumberToReadableString([mod])
						Else

							If num < 1000000000 Then
								[mod] = num Mod 1000000
								i = Int((num - [mod]) / 1000000)

								Select Case i
									Case 1
										result = "unmilione"
									Case Else
										result = convertNumberToReadableString(i) & "milioni"
								End Select

								result = result & convertNumberToReadableString([mod])
							Else

								If num < 1000000000000 Then
									[mod] = num Mod 1000000000
									i = Int((num - [mod]) / 1000000000)

									Select Case i
										Case 1
											result = "unmiliardo"
										Case Else
											result = convertNumberToReadableString(i) & "miliardi"
									End Select

									result = result & convertNumberToReadableString([mod])
								End If
							End If
						End If
					End If
				End If
			End If
		End If

		Return result
	End Function

	Public Function RitornaValoreRandom(Massimo As Integer) As Integer
		Static rnd1 As New Random()

		Return rnd1.Next(Massimo)
	End Function

	Public Function generaPassRandom() As String
		Dim chiaveMaiuscole As String = "ABCDEFGHIJKLMNOPQRSTUVZ"
		Dim chiaveMinuscole As String = "abcdefghijklmnopqrstuvz"
		Dim chiaveNumeri As String = "0123456789"
		Dim chiaveSpeciali As String = "!$%/()=?^"
		Dim nuovaPass As String = ""

		Dim c As Integer = RitornaValoreRandom(chiaveMaiuscole.Length - 1) + 1
		nuovaPass &= Mid(chiaveMaiuscole, c, 1)

		For i As Integer = 1 To 5
			c = RitornaValoreRandom(chiaveMinuscole.Length - 1) + 1
			nuovaPass &= Mid(chiaveMinuscole, c, 1)
		Next

		For i As Integer = 1 To 3
			c = RitornaValoreRandom(chiaveNumeri.Length - 1) + 1
			nuovaPass &= Mid(chiaveNumeri, c, 1)
		Next

		c = RitornaValoreRandom(chiaveSpeciali.Length - 1) + 1
		nuovaPass &= Mid(chiaveSpeciali, c, 1)

		Dim wrapper As New CryptEncrypt(CryptPasswordString)
		Dim nuovaPassCrypt As String = wrapper.EncryptData(nuovaPass)

		Return nuovaPass & ";" & nuovaPassCrypt
	End Function

	Public Function PulisceCartellaTemporanea(MP As String) As String
		'Dim thread As New Thread(AddressOf PulisceCartellaTempThread)
		'thread.Start()
		PulisceCartellaTempThread(MP)

		Return "1"
	End Function

	Private Sub PulisceCartellaTempThread(MP As String)
		Dim Quanti As Integer = 0
		Dim gf As New GestioneFilesDirectory
		Dim pp As String = gf.LeggeFileIntero(MP & "\Impostazioni\Paths.txt")
		pp = pp.Trim()
		pp = pp.Replace(vbCrLf, "")
		If Strings.Right(pp, 1) = "\" Then
			pp = Mid(pp, 1, pp.Length - 1)
		End If
		gf.ScansionaDirectorySingola(pp & "\Appoggio")
		Dim Filetti() As String = gf.RitornaFilesRilevati
		Dim qFiletti As String = gf.RitornaQuantiFilesRilevati

		For i As Integer = 1 To qFiletti
			If TipoPATH <> "SQLSERVER" Then
				Filetti(i) = Filetti(i).Replace("\", "/")
				Filetti(i) = Filetti(i).Replace("//", "/")
			End If

			Dim DataFile As DateTime = FileDateTime(Filetti(i))
			Dim Differenza As Integer = DateAndTime.DateDiff(DateInterval.Second, DataFile, Now)
			If Differenza > 30 Then
				File.Delete(Filetti(i))
				Quanti += 1
			End If
		Next

		' Return Quanti
	End Sub

	Public Function SistemaPercorso(pathPassato As String) As String
		Dim pp As String = pathPassato

		pp = pp.Replace(vbCrLf, "").Trim()
		If Strings.Right(pp, 1) = "\" Or Strings.Right(pp, 1) = "/" Then
			pp = Mid(pp, 1, pp.Length - 1)
		End If

		Return pp
	End Function

	Public Sub ScriveLog(MP As String, Squadra As String, NomeFile As String, Cosa As String)
		Dim gf As New GestioneFilesDirectory

		'If nomeFileLogMail = "" Then
		Dim pp As String = gf.LeggeFileIntero(MP & "\Impostazioni\Paths.txt")
		pp = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")

		Dim paths() As String = pp.Split(";")
		Dim PathLog As String = SistemaPercorso(paths(1))
		Dim ppp As String = gf.LeggeFileIntero(MP & "\Impostazioni\PercorsoSito.txt")

		Dim nomeFileLog As String = PathLog & "\" & Squadra & "\" & NomeFile & ".txt"
		gf.CreaDirectoryDaPercorso(nomeFileLog)
		'End If

		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

		gf.ApreFileDiTestoPerScrittura(nomeFileLog)
		gf.ScriveTestoSuFileAperto(Datella & " " & Cosa)
		gf.ChiudeFileDiTestoDopoScrittura()

		gf = Nothing
	End Sub

	Public Function ConverteInByte(Occupazione2 As Double) As FormatoByte
		Dim Occupazione As Double = Occupazione2
		Dim Cosa As String = ""
		Dim giga As Double = 1024L * 1024 * 1024 ' * 1024
		Dim mega As Double = 1024L * 1024 '* 1024
		Dim kappa As Double = 1024L '* 1024

		If Occupazione > giga Then
			Occupazione /= giga
			Cosa = "Gb."
		Else
			If Occupazione > mega Then
				Occupazione /= mega
				Cosa = "Mb."
			Else
				If Occupazione > kappa Then
					Occupazione /= kappa
					Cosa = "Kb."
				Else
					Cosa = "B."
				End If
			End If
		End If
		Occupazione = CInt(Occupazione * 100) / 100

		Dim Ritorno As New FormatoByte
		Ritorno.Occupazione = Occupazione
		Ritorno.Cosa = Cosa

		Return Ritorno
	End Function

	Public Function GeneraRicevutaEScontrino(MP As String, Squadra As String, NomeSquadra As String, idAnno As String, idGiocatore As String, idPagamento As String, idUtente As String, vecchioID As String) As String
		Dim Ritorno As String = ""
		Dim Ok As Boolean = True

		Try
			Dim Connessione As String = LeggeImpostazioniDiBase(MP, Squadra)

			If Connessione = "" Then
				Ritorno = ErroreConnessioneNonValida
			Else
				Dim Conn As Object = New clsGestioneDB(Squadra)

				If TypeOf (Conn) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
				Else
					ScriveLog(MP, Squadra, "Pagamenti", "ID Giocatore: " & idGiocatore & " - ID Pagamento: " & idPagamento)
					Dim Rec As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
					Dim Sql As String = "Select * From GiocatoriPagamenti Where idGiocatore=" & idGiocatore & " And Progressivo=" & vecchioID
					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Dati ricevuta non presenti"
						ScriveLog(MP, Squadra, "Pagamenti", "Dati ricevuta NON acquisiti")
					Else
						Dim Pagamento As String = "" & Rec("Pagamento").Value
						Dim DataRicevuta As String = "" & Rec("DataPagamento").Value
						Dim Commento As String = "" & Rec("Commento").Value
						Dim idPagatore As String = "" & Rec("idUtentePagatore").Value
						Dim idRegistratore As String = "" & Rec("idUtenteRegistratore").Value
						Dim Note As String = "" & Rec("Note").Value
						Dim Validato As String = "" & Rec("Validato").Value
						Dim idTipoPagamento As String = "" & Rec("idTipoPagamento").Value
						Dim idRata As String = "" & Rec("idRata").Value
						Dim idQuota As String = "" & Rec("idQuota").Value
						Dim idModalitaPagamento As String = "" & Rec("MetodoPagamento").Value
						Dim NumeroRicevuta As String = "" & Rec("NumeroRicevuta").Value
						'Rec.Close()

						ScriveLog(MP, Squadra, "Pagamenti", "Dati ricevuta acquisiti: Pagamento " & Pagamento)

						Dim nomeRate As String = ""
						Dim rr As New List(Of String)
						If idRata.Contains(";") Then
							Dim r() As String = idRata.Split(";")
							For Each r2 As String In r
								rr.Add(r2)
							Next
						Else
							Dim r As String = idRata
							rr.Add(r)
						End If

						For Each r As String In rr
							If r <> "" Then
								Sql = "Select * From QuoteRate Where idQuota=" & idQuota & " And Progressivo=" & r
								Rec = Conn.LeggeQuery(MP, Sql, Connessione)
								If Not Rec.Eof() Then
									nomeRate &= ("" & Rec("DescRata").Value) & "<br />"
								End If
							End If
						Next

						ScriveLog(MP, Squadra, "Pagamenti", "Nome rate " & nomeRate)

						Dim Cognome As String = ""
						Dim CognomePagatore As String = ""
						Dim Nome As String = ""
						Dim CognomeIscritto As String = ""
						Dim NomeIscritto As String = ""
						Dim CodFiscalePagatore As String = ""
						Dim CodFiscaleIscritto As String = ""
						Dim NomePolisportiva As String = ""
						Dim Indirizzo As String = ""
						Dim CodiceFiscale As String = ""
						Dim PIva As String = ""
						Dim Telefono As String = ""
						Dim eMail As String = ""
						Dim indirizzoPagatore As String = ""
						Dim Suffisso As String = ""

						Sql = "SELECT * FROM Anni"
						Rec = Conn.LeggeQuery(MP, Sql, Connessione)
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessuna squadra rilevata"
							ScriveLog(MP, Squadra, "Pagamenti", Ritorno)
							Ok = False
						Else
							NomeSquadra = "" & Rec("NomeSquadra").Value
							NomePolisportiva = "" & Rec("NomePolisportiva").Value
							Indirizzo = "" & Rec("Indirizzo").Value
							CodiceFiscale = "" & Rec("CodiceFiscale").Value
							PIva = "" & Rec("PIva").Value
							Telefono = "" & Rec("Telefono").Value
							eMail = "" & Rec("Mail").Value
							Suffisso = "" & Rec("Suffisso").Value
						End If
						'Rec.Close()

						ScriveLog(MP, Squadra, "Pagamenti", "Dati squadra acquisiti: " & NomeSquadra)

						If Ok Then
							If idPagatore = 3 Then
								Sql = "SELECT * FROM Giocatori Where idGiocatore=" & idGiocatore
								Rec = Conn.LeggeQuery(MP, Sql, Connessione)
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
									ScriveLog(MP, Squadra, "Pagamenti", Ritorno)
									Ok = False
								Else
									Cognome = Rec("Cognome").Value
									Nome = Rec("Nome").Value
									CodFiscalePagatore = Rec("CodFiscale").Value

									CognomeIscritto = Rec("Cognome").Value
									NomeIscritto = Rec("Nome").Value
									CodFiscaleIscritto = Rec("CodFiscale").Value
								End If
								Rec.Close()
							Else
								Sql = "SELECT * FROM Giocatori Where idGiocatore=" & idGiocatore
								Rec = Conn.LeggeQuery(MP, Sql, Connessione)
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
									ScriveLog(MP, Squadra, "Pagamenti", Ritorno)
									Ok = False
								Else
									CognomeIscritto = Rec("Cognome").Value
									NomeIscritto = Rec("Nome").Value
									CodFiscaleIscritto = Rec("CodFiscale").Value
								End If
								Rec.Close()

								Sql = "SELECT * FROM GiocatoriDettaglio Where idGiocatore=" & idGiocatore
								Rec = Conn.LeggeQuery(MP, Sql, Connessione)
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
									ScriveLog(MP, Squadra, "Pagamenti", Ritorno)
									Ok = False
								Else
									If idPagatore = 1 Then
										CognomePagatore = "" & Rec("Genitore1").Value
										indirizzoPagatore = "" & Rec("Indirizzo1").Value
										CodFiscalePagatore = "" & Rec("CodFiscale1").Value
									Else
										CognomePagatore = "" & Rec("Genitore2").Value
										indirizzoPagatore = "" & Rec("Indirizzo2").Value
										CodFiscalePagatore = "" & Rec("CodFiscale2").Value
									End If
								End If
								Rec.Close()
							End If
						End If

						ScriveLog(MP, Squadra, "Pagamenti", "Dati pagatore acquisiti: " & Cognome & " " & Nome)

						If Ok Then
							Dim Intero As String
							Dim Virgola As String

							If Pagamento.Contains(",") Or Pagamento.Contains(".") Then
								If Pagamento.Contains(".") Then
									Dim pp1() As String = Pagamento.Split(".")
									Intero = pp1(0)
									Virgola = pp1(1)
								Else
									Dim pp22() As String = Pagamento.Split(",")
									Intero = pp22(0)
									Virgola = pp22(1)
								End If
							Else
								Intero = Pagamento
								Virgola = ""
							End If

							If Virgola = "" Then
								Virgola = "00"
							Else
								If Virgola.Length = 1 Then
									Virgola = "0" & Virgola
								Else
									If Virgola > 2 Then
										Virgola = Mid(Virgola, 1, 2)
									End If
								End If
							End If

							Dim Cifre1 As String = convertNumberToReadableString(Val(Intero))
							Dim Cifre2 As String = convertNumberToReadableString(Val(Virgola))
							Dim Altro2 As String = ""
							If Cifre2 <> "" Then
								Altro2 = "/" & Virgola
							End If
							Dim ImportoLettere As String = Cifre1 & Altro2

							Dim Dati As String = "C.F.: " & CodiceFiscale & " P.I.:" & PIva & "<br />Telefono: " & Telefono & "<br />E-Mail: " & eMail
							Dim Altro As String = ""
							If Commento <> "" Then
								Altro = "- " & Commento
							End If
							Dim sDataRicevuta As String = ""
							If DataRicevuta <> "" Then
								Dim d() As String = DataRicevuta.Split("-")
								sDataRicevuta = d(2) & "/" & d(1) & "/" & d(0)
							Else
								sDataRicevuta = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year
							End If
							Dim Motivazione As String = CognomeIscritto & " " & NomeIscritto & " " & CodFiscaleIscritto & " " & Altro
							Dim ssNumeroRicevuta As String = ""
							If NumeroRicevuta <> "" Then
								ssNumeroRicevuta = NumeroRicevuta
							Else
								If Suffisso <> "" Then
									ssNumeroRicevuta = NumeroRicevuta & "/" & Suffisso & "/" & Now.Year
								Else
									ssNumeroRicevuta = NumeroRicevuta & "/" & Now.Year
								End If
							End If
							Dim NominativoRicevuta As String = CognomeIscritto & " " & NomeIscritto & "<br />" & CodFiscaleIscritto & " " & Altro & "<br />" & nomeRate

							ScriveLog(MP, Squadra, "Pagamenti", "Nominativo ricevuta: " & NominativoRicevuta)

							Dim gT1 As New GestioneTags(MP)
							ScriveLog(MP, Squadra, "Pagamenti", "Eseguo stampa ricevuta")
							gT1.EsegueStampaRicevuta(MP, Squadra, idGiocatore, idAnno, idPagamento, Dati, ssNumeroRicevuta, sDataRicevuta, Motivazione, Intero, Virgola, ImportoLettere, NominativoRicevuta, idPagatore)
							gT1 = Nothing

							Dim gT2 As New GestioneTags(MP)
							ScriveLog(MP, Squadra, "Pagamenti", "Eseguo stampa scontrino")
							gT2.EsegueStampaScontrino(MP, Squadra, idGiocatore, idAnno, idPagamento, Dati, ssNumeroRicevuta, sDataRicevuta, Motivazione, Intero, Virgola, ImportoLettere, NominativoRicevuta, idPagatore)
							gT2 = Nothing

							Ritorno = "*"

							'Dim gf As New GestioneFilesDirectory
							'Dim filePaths As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")
							'Dim p() As String = filePaths.Split(";")
							'If Strings.Right(p(0), 1) <> "\" Then
							'	p(0) &= "\"
							'End If
							'p(2) = p(2).Replace(vbCrLf, "").Trim
							'If Strings.Right(p(2), 1) <> "/" Then
							'	p(2) = p(2) & "/"
							'End If
							'' Dim url As String = p(2) & NomeSquadra.Replace(" ", "_") & "/Societa/" & idAnno & "_1.jpg"

							'Dim pp As String = gf.LeggeFileIntero(MP & "\Impostazioni\Paths.txt")
							'pp = pp.Replace(vbCrLf, "").Trim
							'If Strings.Right(pp, 1) = "\" Then
							'	pp = Mid(pp, 1, pp.Length - 1)
							'End If
							'Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

							'Dim nomeImm As String = p(2) & NomeSquadra.Replace(" ", "_") & "/Societa/" & idAnno & "_1.kgb"
							'Dim pathImm As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\" & idAnno & "_1.kgb"
							'Dim nomeImmConv As String = ""
							'Dim c As New CriptaFiles
							'If ControllaEsistenzaFile(pathImm) Then
							'	nomeImmConv = p(2) & "" & NomeSquadra.Replace(" ", "_") & "/Societa/Societa_1.png"
							'	Dim pathImmConv As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\Societa_1.png"
							'	c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)
							'End If

							'Dim pathRicevuta As String = p(0) & Squadra & "\Scheletri\ricevuta_pagamento.txt"
							'If Not ControllaEsistenzaFile(pathRicevuta) Then
							'	pathRicevuta = MP & "\Scheletri\ricevuta_pagamento.txt"
							'End If
							'Dim Body As String = gf.LeggeFileIntero(pathRicevuta)
							'Dim path As String = p(0) & "\" & Squadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\"
							'gf.CreaDirectoryDaPercorso(path)
							'Dim fileFinale As String = path & "Ricevuta_" & idPagamento & ".pdf"
							'Dim fileAppoggio As String = path & "Ricevuta_" & idPagamento & ".html"

							'If vecchioID <> "-1" Then
							'	Dim fileFinaleVecchio As String = path & "Ricevuta_" & vecchioID & ".pdf"
							'	Dim fileAppoggioVecchio As String = path & "Ricevuta_" & vecchioID & ".html"
							'	Try
							'		File.Delete(fileFinaleVecchio)
							'	Catch ex As Exception

							'	End Try
							'	Try
							'		File.Delete(fileAppoggioVecchio)
							'	Catch ex As Exception

							'	End Try
							'End If

							'Dim Intero As String
							'Dim Virgola As String

							'If Pagamento.Contains(",") Or Pagamento.Contains(".") Then
							'	If Pagamento.Contains(".") Then
							'		Dim pp1() As String = Pagamento.Split(".")
							'		Intero = pp1(0)
							'		Virgola = pp1(1)
							'	Else
							'		Dim pp22() As String = Pagamento.Split(",")
							'		Intero = pp22(0)
							'		Virgola = pp22(1)
							'	End If
							'Else
							'	Intero = Pagamento
							'	Virgola = ""
							'End If

							'If Virgola = "" Then
							'	Virgola = "00"
							'Else
							'	If Virgola.Length = 1 Then
							'		Virgola = "0" & Virgola
							'	Else
							'		If Virgola > 2 Then
							'			Virgola = Mid(Virgola, 1, 2)
							'		End If
							'	End If
							'End If

							'Dim Dati As String = "C.F.: " & CodiceFiscale & " P.I.:" & PIva & "<br />Telefono: " & Telefono & "<br />E-Mail: " & eMail
							'Dim Altro As String = ""
							'If Commento <> "" Then
							'	Altro = "- " & Commento
							'End If

							'Body = Body.Replace("***URL LOGO***", nomeImmConv)
							'Body = Body.Replace("***NOME POLISPORTIVA***", NomePolisportiva)
							'Body = Body.Replace("***INDIRIZZO***", Indirizzo)
							'Body = Body.Replace("***DATI***", Dati)
							'If NumeroRicevuta <> "" Then
							'	Body = Body.Replace("***NUMERO_RICEVUTA***", NumeroRicevuta)
							'Else
							'	If Suffisso <> "" Then
							'		Body = Body.Replace("***NUMERO_RICEVUTA***", idPagamento & "/" & Suffisso & "/" & Now.Year)
							'	Else
							'		Body = Body.Replace("***NUMERO_RICEVUTA***", idPagamento & "/" & Now.Year)
							'	End If
							'End If
							'If DataRicevuta <> "" Then
							'	Dim d() As String = DataRicevuta.Split("-")
							'	Dim sDataRicevuta As String = d(2) & "/" & d(1) & "/" & d(0)
							'	Body = Body.Replace("***DATA_RICEVUTA***", sDataRicevuta) ' Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
							'Else
							'	Body = Body.Replace("***DATA_RICEVUTA***", Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
							'End If
							'Body = Body.Replace("***NOME***", CognomePagatore & "<br />" & CodFiscalePagatore & "<br />" & indirizzoPagatore)
							'Body = Body.Replace("***MOTIVAZIONE***", CognomeIscritto & " " & NomeIscritto & "<br />" & CodFiscaleIscritto & " " & Altro & "<br />" & nomeRate)
							'Body = Body.Replace("***IMPORTO***", Intero)
							'Body = Body.Replace("***VIRGOLE***", Virgola)

							'Dim Cifre1 As String = convertNumberToReadableString(Val(Intero))
							'Dim Cifre2 As String = convertNumberToReadableString(Val(Virgola))
							'Dim Altro2 As String = ""
							'If Cifre2 <> "" Then
							'	Altro2 = "/" & Virgola
							'End If
							'Body = Body.Replace("***IMPORTO LETTERE***", Cifre1 & Altro2)

							'filePaths = gf.LeggeFileIntero(MP & "\Impostazioni\Paths.txt")
							'filePaths = filePaths.Replace(vbCrLf, "").Trim
							'If Strings.Right(filePaths, 1) <> "\" Then
							'	filePaths &= "\"
							'End If
							'' Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & idGiocatore & "_" & idPagatore & ".png"
							'' Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Segreteria\" & idAnno & ".kgb"

							'Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_").Trim & "\Utenti\" & idAnno & "_" & idUtente & "_Firma.kgb"
							''Sql = "rollback"
							''Dim Ritorno2 As String = Conn.EsegueSql(MP, Sql, Connessione)
							''Return pathFirma
							'If ControllaEsistenzaFile(pathFirma) Then
							'	Dim urlFirma As String = pp & "\" & NomeSquadra.Replace(" ", "_").Trim & "\Utenti\" & idAnno & "_" & idUtente & "_Firma.kgb"
							'	'Dim pathFirmaConv As String = p(2) & "/Appoggio/Firma_" & Esten & ".png"
							'	Dim urlFirmaConv As String = pp & "\Appoggio\Firma_" & Esten & ".png"
							'	c.DecryptFile(CryptPasswordString, urlFirma, urlFirmaConv)

							'	Body = Body.Replace("***URL FIRMA***", urlFirmaConv)
							'Else
							'	Body = Body.Replace("***URL FIRMA***", "")
							'End If

							'' Body = Body & "<hr /><div style=""text-algin: center; width: 100%;"">Stampato tramite InCalcio – www.incalcio.it – info@incalcio.it</div>"

							'gf.EliminaFileFisico(fileAppoggio)
							'gf.ApreFileDiTestoPerScrittura(fileAppoggio)
							'gf.ScriveTestoSuFileAperto(Body)

							'gf.ChiudeFileDiTestoDopoScrittura()

							'' Scontrino
							'Dim pathScontr As String = p(0) & Squadra & "\Scheletri\ricevuta_scontrino.txt"
							'If Not ControllaEsistenzaFile(pathScontr) Then
							'	pathScontr = MP & "\Scheletri\ricevuta_scontrino.txt"
							'End If
							'Dim BodyScontrino As String = gf.LeggeFileIntero(pathScontr)
							'Dim pathScontrino As String = p(0) & "\" & Squadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\"
							'gf.CreaDirectoryDaPercorso(pathScontrino)
							'Dim fileFinaleScontrino As String = path & "Scontrino_" & idPagamento & ".pdf"
							'Dim fileAppoggioScontrino As String = path & "Scontrino_" & idPagamento & ".html"

							'If vecchioID <> "-1" Then
							'	Dim fileFinaleScontrinoVecchio As String = path & "Scontrino_" & vecchioID & ".pdf"
							'	Dim fileAppoggioScontrinoVecchio As String = path & "Scontrino_" & vecchioID & ".html"
							'	Try
							'		File.Delete(fileFinaleScontrinoVecchio)
							'	Catch ex As Exception

							'	End Try
							'	Try
							'		File.Delete(fileAppoggioScontrinoVecchio)
							'	Catch ex As Exception

							'	End Try
							'End If

							'BodyScontrino = BodyScontrino.Replace("***NOME POLISPORTIVA***", NomePolisportiva)
							'BodyScontrino = BodyScontrino.Replace("***INDIRIZZO***", Indirizzo)
							'BodyScontrino = BodyScontrino.Replace("***DATI***", Dati)
							'If NumeroRicevuta <> "" Then
							'	BodyScontrino = BodyScontrino.Replace("***NUMERO_RICEVUTA***", NumeroRicevuta)
							'Else
							'	If Suffisso <> "" Then
							'		BodyScontrino = BodyScontrino.Replace("***NUMERO_RICEVUTA***", idPagamento & "/" & Suffisso & "/" & Now.Year)
							'	Else
							'		BodyScontrino = BodyScontrino.Replace("***NUMERO_RICEVUTA***", idPagamento & "/" & Now.Year)
							'	End If
							'End If
							'If DataRicevuta <> "" Then
							'	Dim d() As String = DataRicevuta.Split("-")
							'	Dim sDataRicevuta As String = d(2) & "/" & d(1) & "/" & d(0)
							'	BodyScontrino = BodyScontrino.Replace("***DATA_RICEVUTA***", sDataRicevuta) ' Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
							'Else
							'	BodyScontrino = BodyScontrino.Replace("***DATA_RICEVUTA***", Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
							'End If
							'BodyScontrino = BodyScontrino.Replace("***MOTIVAZIONE***", CognomeIscritto & " " & NomeIscritto & "<br />" & CodFiscaleIscritto & "<br />" & Altro & "<br />" & nomeRate)
							'BodyScontrino = BodyScontrino.Replace("***IMPORTO***", Intero & "." & Virgola)

							'nomeImm = p(2) & NomeSquadra.Replace(" ", "_") & "/Societa/" & idAnno & "_1.kgb"
							'pathImm = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\" & idAnno & "_1.kgb"
							'If ControllaEsistenzaFile(pathImm) Then
							'	nomeImmConv = p(2) & "/" & NomeSquadra.Replace(" ", "_") & "/Societa/Societa_1.png"
							'	Dim pathImmConv As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\Societa_1.png"
							'	c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)

							'	BodyScontrino = BodyScontrino.Replace("***immagine logo menu settaggi***", "<img src=""" & nomeImmConv & """ style=""width: 240px; height: 240px;"" />")
							'Else
							'	BodyScontrino = BodyScontrino.Replace("***immagine logo menu settaggi***", "")
							'End If
							'BodyScontrino = BodyScontrino.Replace("***NOME***", CognomePagatore & " " & indirizzoPagatore & "<br />" & CodFiscalePagatore)

							'BodyScontrino = BodyScontrino & "<hr /><div style=""text-algin: center; width: 100%;"">Stampato tramite InCalcio – www.incalcio.it<br />info@incalcio.it</div>"

							'gf.EliminaFileFisico(fileAppoggioScontrino)
							'gf.ApreFileDiTestoPerScrittura(fileAppoggioScontrino)
							'gf.ScriveTestoSuFileAperto(BodyScontrino)
							'gf.ChiudeFileDiTestoDopoScrittura()
							'' Scontrino

							'Dim pp2 As New pdfGest
							'Ritorno = pp2.ConverteHTMLInPDF(fileAppoggio, fileFinale, "")
							'Dim Ritorno2 As String = pp2.ConverteHTMLInPDF(fileAppoggioScontrino, fileFinaleScontrino, "", True)
							'If Ritorno<> "OK" And Ritorno2<> "OK" Then
							'	Ok = False
							'Else
							'	If Ritorno2<> "OK" Then
							'		Ritorno = Ritorno2
							'	End If
							'End If
						End If
					End If
				End If
			End If

		Catch ex As Exception
			Ritorno = StringaErrore & " " & ex.Message
		End Try

		Return Ritorno
	End Function

	'Public Function DecriptaImmagine(NomeSquadra As String, Tipologia As String, NomeImmagine As String) As String
	'Dim c As New CriptaFiles
	'Dim Immagine As String = ""
	'Dim Esten2 As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)
	'Dim pathImmagine As String = P(2) & "/" & NomeSquadra.Replace(" ", "_") & "/" & Tipologia & "/" & NomeImmagine & ".kgb"
	'Dim urlImmagine As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\" & Tipologia & "\" & NomeImmagine & ".kgb"
	'Dim pathImmagineConvertita As String = P(2) & "/Appoggio/" & NomeImmagine & "_" & Esten2 & ".png"
	'Dim urlImmagineConvertita As String = pp & "\Appoggio\" & NomeImmagine & "_" & Esten2 & ".png"
	'If ControllaEsistenzaFile(urlImmagine) Then
	'	c.DecryptFile(CryptPasswordString, pathImmagine, pathImmagineConvertita)

	'	Immagine = "<img src=""" & urlImmagineConvertita & """ style=""width: 50px; height: 50px;"" />"
	'Else
	'	Immagine = ""
	'End If

	'Return Immagine
	'End Function

	Public Function ConverteData(Data As String) As String
		Dim Ritorno As String = Data
		If Ritorno <> "" And Ritorno.Contains("-") Then
			Dim dd() As String = Ritorno.Split("-")
			' Ritorno = Format(Val(dd(2)), "00") & "-" & Format(dd(1), "00") & "-" & dd(0)
			If dd(2) > 50 Then
				Ritorno = dd(0) & "-" & dd(1) & "-" & dd(2)
			Else
				Ritorno = dd(2) & "-" & dd(1) & "-" & dd(0)
			End If
		End If

		Return Ritorno
	End Function

	Public Function CreaNumeroTesseraNFC(MP As String, Conn As Object, Connessione As String, Squadra As String, idGiocatore As String) As String
		Dim CodiceTessera As String = ""
		Dim Rec As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
		Dim Sql As String = "Select * From [Generale].[dbo].[GiocatoriTessereNFC] Where idGiocatore=" & idGiocatore & " And CodSquadra='" & Squadra & "'"
		Rec = Conn.LeggeQuery(MP, Sql, Connessione)
		If Rec.Eof() Then
			CodiceTessera = DateTime.Now.Year & Strings.Format(DateTime.Now.Month, "00") & Strings.Format(DateTime.Now.Day, "00") & Strings.Format(DateTime.Now.Hour, "00") & Strings.Format(DateTime.Now.Minute, "00") + Strings.Format(DateTime.Now.Second, "00")
			Dim stringaRandom As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
			Dim r As String = ""
			For i As Integer = 1 To 6
				Dim p As String = RitornaValoreRandom(stringaRandom.Length - 1) + 1
				r &= Mid(stringaRandom, p, 1)
			Next
			CodiceTessera &= r
			Sql = "Insert Into [Generale].[dbo].[GiocatoriTessereNFC] Values (" & idGiocatore & ", '" & Squadra & "', '" & CodiceTessera & "')"
			Dim Ritorno As String = Conn.EsegueSql(MP, Sql, Connessione)
		End If
		Rec.Close()

		Return CodiceTessera
	End Function

	Public Function RitornaCategorieUtente(MP As String, Conn As Object, Connessione As String, Utente As String) As String
		Dim Ritorno As String = ""
		Dim Rec As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
		Dim Sql As String = ""

		Dim idGiocatore As String = ""

		Sql = "Select * From [Generale].[dbo].[Utenti] Where EMail = '" & Utente.Replace("'", "''") & "'"
		Rec = Conn.LeggeQuery(MP, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Dim idTipologia As String = ""

			If Not Rec.Eof() Then
				idTipologia = Rec("idTipologia").Value
				idGiocatore = Rec("idGiocatore").Value
			Else
				Rec.Close()

				' Forse è un giocatore ?
				Sql = "Select * From Giocatori Where EMail = '" & Utente.Replace("'", "''") & "'"
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						idTipologia = 6
						idGiocatore = Rec("idGiocatore").Value
					End If

					Rec.Close()
				End If
			End If

			Select Case idTipologia
				Case "0"
					' Super User
					Ritorno = "-1"
				Case "1"
					' Amministratore
					Ritorno = "-1"
				Case "2"
					' Utente
				Case "3"
					' Genitore
					If idGiocatore <> "" Then
						Dim idG() As String = idGiocatore.Split(";")
						Dim Ricerca As String = ""

						For Each idC As String In idG
							If idC <> "" Then
								Ricerca &= idC & ","
							End If
						Next

						If Ricerca <> "" Then
							Ricerca = "(" & Mid(Ricerca, 1, Ricerca.Length - 1) & ")"

							Sql = "Select Categorie From Giocatori Where idGiocatore In " & Ricerca
							Rec = Conn.LeggeQuery(MP, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								Dim Cat As String = ""

								Do Until Rec.Eof()
									Cat &= Rec("Categorie").Value

									Rec.MoveNext
								Loop
								Rec.Close()

								If Cat <> "" Then
									Dim Cat2() As String = Cat.Split("-")
									Dim Cate As New List(Of String)

									For Each c As String In Cat2
										Dim Ok As Boolean = True

										For Each cc As String In Cate
											If cc = c Then
												Ok = False
												Exit For
											End If
										Next

										If Ok Then
											Cate.Add(c)
										End If
									Next

									For Each cate2 As String In Cate
										Ritorno &= cate2 & ";"
									Next
								End If
							End If
						End If
					End If
				Case "4", "8"
					' Dirigente
					Sql = "Select B.idCategoria From Dirigenti A " &
						"Left Join DirigentiCategorie B On A.idDirigente = B.idUtente " &
						"Where A.EMail = '" & Utente.Replace("'", "''") & "' And A.Eliminato = 'N'"
					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof()
							Ritorno &= Rec("idCategoria").Value & ";"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Case "6"
					' Giocatore
					Sql = "Select Categorie From Giocatori Where idGiocatore In " & idGiocatore
					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof() Then
							Dim Cat2() As String = Rec("Categorie").Value.Split("-")
							Dim Cate As New List(Of String)

							For Each c As String In Cat2
								Dim Ok As Boolean = True

								For Each cc As String In Cate
									If cc = c Then
										Ok = False
										Exit For
									End If
								Next

								If Ok Then
									Cate.Add(c)
								End If
							Next

							For Each cate2 As String In Cate
								Ritorno &= cate2 & ";"
							Next
						End If

						Rec.Close()
					End If
				Case "5", "7"
					' Allenatore
					Sql = "Select B.idCategoria From Allenatori A " &
						"Left Join AllenatoriCategorie B On A.idAllenatore = B.idUtente " &
						"Where A.EMail = '" & Utente.Replace("'", "''") & "' And A.Eliminato = 'N'"
					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof()
							Ritorno &= Rec("idCategoria").Value & ";"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
			End Select
		End If

		Return Ritorno
	End Function

	Public Function ControllaEsistenzaFile(Filetto As String) As Boolean
		Dim Ritorno As Boolean = True

		If TipoPATH <> "SQLSERVER" Then
			Filetto = Filetto.Replace("\", "/")
			Filetto = Filetto.Replace("//", "/")
		End If

		If File.Exists(Filetto) Then
			Return True
		Else
			Return False
		End If
	End Function

	Public Function RitornaImmagine(MP As String, Tipologia As String, Squadra As String, id As String) As String
		Dim Ritorno As String = ""
		Dim Rec As Object

		Dim Connessione As String = LeggeImpostazioniDiBase(MP, Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Sql = "Select * From immagini_" & Tipologia & " Where id=" & id
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = ""
					Else
						Ritorno = Rec("Dati").Value
					End If
					Rec.Close()
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function ControllaEsistenzaFile(MP As String, Conn As Object, Connessione As String, idGiocatore As String, Genitore As String, Privacy As String) As Boolean
		Dim Ritorno As Boolean = True
		Dim Rec As Object

		Dim Sql As String = "Select * From immagini_firme Where id=" & idGiocatore & " And Progressivo=" & Genitore & " And Privacy='" & Privacy.ToLower & "'"
		Rec = Conn.LeggeQuery(MP, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof() Then
				Ritorno = False
			Else
				Ritorno = True
			End If
			Rec.Close()
		End If

		Return Ritorno
	End Function
End Module
