Imports System.Reflection
Imports System.Timers
Imports ADODB
Imports MySqlConnector

Public Class clsGestioneDB
	Private Structure LogStruct
		Dim Cosa As String
		Dim Dove As String
	End Structure

	Private mdb As clsMariaDB
	Private nomeFileLogQuery As String = ""
	Private nomeFileLogExec As String = ""
	Private listaLog As New List(Of LogStruct)
	Private timerLog As Timers.Timer = Nothing
	Private Squadra As String = ""

	Public Sub New(SQ As String)
		Squadra = SQ
	End Sub

	Public Function ApreDB(ByVal Connessione As String) As Object
		' Routine che apre il DB e vede se ci sono errori
		Dim Conn As Object
		Dim TipoDB As String = LeggeTipoDB()

		If TipoDB = "SQLSERVER" Then
			Conn = CreateObject("ADODB.Connection")
			Try
				Conn.Open(Connessione)
				Conn.CommandTimeout = 0
			Catch ex As Exception
				Conn = StringaErrore & " " & ex.Message
			End Try
		Else
			mdb = New clsMariaDB

			Try
				Conn = mdb.apreConnessione(Connessione)
			Catch ex As Exception
				Conn = StringaErrore & " " & ex.Message
			End Try
		End If

		Return Conn
	End Function

	Public Function EsegueSql(MP As String, Sql As String, Connessione As String, Optional ModificaQuery As Boolean = True) As String
		Dim Ritorno As String = "*"
		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
		Dim gf As New GestioneFilesDirectory
		Dim Conn As Object = ApreDB(Connessione)
		Dim Sql2 As String = ""

		If ModificaQuery Then
			If TipoDB = "SQLSERVER" Then
				Sql2 = Sql
			Else
				Sql2 = ConverteStringaPerLinux(Sql)
			End If
		Else
			Sql2 = Sql
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			'If nomeFileLogGenerale = "" Then
			Dim paths As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")
			Dim pp() As String = paths.Split(";")
			pp(1) = pp(1).Replace(vbCrLf, "")
			If Strings.Right(pp(1), 1) <> "\" Then
				pp(1) = pp(1) & "\"
			End If
			If Squadra = "" Then
				nomeFileLogExec = pp(1) & "Generale\Exec_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
			Else
				nomeFileLogExec = pp(1) & Squadra & "\Exec_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
			End If

			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogExec)

			Dim sql3 As String = Sql2
			If sql3.Length > 1024 And (sql3.ToUpper.Contains("IMMAGINI_") Or sql3.ToUpper.Contains("ALLEGATI_")) Then
				sql3 = Mid(sql3, 1, 700) & "..." & Right(Sql2, 324)
			End If
			ThreadScriveLog(Datella & ": " & sql3, nomeFileLogExec)
			' End If
		End If

		If TypeOf (Conn) Is String Then
			If effettuaLog And Not HttpContext.Current Is Nothing Then
				ThreadScriveLog(Datella & " Errore connessione: " & Conn, nomeFileLogQuery)
				ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogExec)
				ThreadScriveLog("", nomeFileLogExec)
			End If

			Return Conn
		End If

		' Routine che esegue una query sul db
		If TipoDB = "SQLSERVER" Then
			Try
				Conn.Execute(Sql2)
				If effettuaLog Then
					ThreadScriveLog(Datella & ": OK", nomeFileLogExec)
				End If
			Catch ex As Exception
				Ritorno = StringaErrore & " " & ex.Message
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogExec)
				End If

			End Try
		Else
			Try
				Ritorno = mdb.EsegueSql(Sql2, ModificaQuery)
				If Ritorno.ToUpper <> "OK" Then
					Ritorno = StringaErrore & " " & Ritorno
					If effettuaLog Then
						ThreadScriveLog(Datella & ": " & Ritorno, nomeFileLogExec)
					End If
				Else
					If effettuaLog Then
						ThreadScriveLog(Datella & ": OK", nomeFileLogExec)
					End If
				End If
			Catch ex As Exception
				Ritorno = StringaErrore & " " & ex.Message
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogExec)
				End If

			End Try
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogExec)
			ThreadScriveLog("", nomeFileLogExec)
		End If

		ChiudeDB(Conn)

		Return Ritorno
	End Function

	Public Sub Close()

	End Sub

	Public Function LeggeQuery(MP As String, Sql As String, ByVal Connessione As String, Optional ModificaQuery As Boolean = True) As Object
		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
		Dim gf As New GestioneFilesDirectory
		Dim TipoDB As String = LeggeTipoDB()
		Dim Conn As Object = ApreDB(Connessione)
		Dim Sql2 As String = ""

		If ModificaQuery = True Then
			If TipoDB = "SQLSERVER" Then
				Sql2 = Sql
			Else
				Sql2 = ConverteStringaPerLinux(Sql)
			End If
		Else
			Sql2 = Sql
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			'If nomeFileLogGenerale = "" Then
			Dim paths As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")

			Dim pp() As String = paths.Split(";")
			pp(1) = pp(1).Replace(vbCrLf, "")
			If Strings.Right(pp(1), 1) <> "\" Then
				pp(1) = pp(1) & "\"
			End If
			If Squadra = "" Then
				nomeFileLogQuery = pp(1) & "Generale\Query_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
			Else
				nomeFileLogQuery = pp(1) & Squadra & "\Query_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
			End If

			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogQuery)
			ThreadScriveLog(Datella & " Modifica Query: " & ModificaQuery, nomeFileLogQuery)
			ThreadScriveLog(Datella & " TIPO DB: " & TipoDB, nomeFileLogQuery)
			Dim sql3 As String = Sql2
			If sql3.Length > 1024 And (sql3.ToUpper.Contains("IMMAGINI_") Or sql3.ToUpper.Contains("ALLEGATI_")) Then
				sql3 = Mid(sql3, 1, 700) & "..." & Right(Sql2, 324)
			End If
			ThreadScriveLog(Datella & ": " & sql3, nomeFileLogQuery)
			'End If
		End If

		If TypeOf (Conn) Is String Then
			If effettuaLog And Not HttpContext.Current Is Nothing Then
				ThreadScriveLog(Datella & " Errore connessione: " & Conn, nomeFileLogQuery)
				ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogQuery)
				ThreadScriveLog("", nomeFileLogQuery)
			End If

			Return Conn
		End If

		Dim Rec As Object = Nothing

		If TipoDB = "SQLSERVER" Then
			Rec = New Recordset

			Try
				Rec.Open(Sql2, Conn)
			Catch ex As Exception
				Rec = StringaErrore & " " & ex.Message
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogQuery)
				End If
			End Try
		Else
			Try
				Rec = mdb.Lettura(Sql2, ModificaQuery)
			Catch ex As Exception
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogQuery)
				End If
				' Return StringaErrore & " " & ex.Message
			End Try
		End If

		If TypeOf (Rec) Is String Then
			If effettuaLog Then
				ThreadScriveLog(Datella & ": ERRORE SQL -> " & Rec, nomeFileLogQuery)
			End If
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogQuery)
			ThreadScriveLog("", nomeFileLogQuery)
		End If

		ChiudeDB(Conn)

		Return Rec
	End Function

	Private Function ConverteStringaPerLinux(Sql As String) As String
		Dim Sql2 As String = Sql

		If Sql2.ToUpper.Trim.StartsWith("INSERT INTO ") Then
			Dim a As Integer = Sql2.ToUpper.IndexOf(" VALUES")

			If a = 0 Then
				a = Sql2.ToUpper.IndexOf(" SELECT")
			End If
			If a > 0 Then
				Dim inizio As String = Mid(Sql2, 1, a)
				Dim modificato As String = inizio.ToLower
				Sql2 = Sql2.Replace(inizio, modificato)
			End If
		Else
			If Sql2.ToUpper.Trim.StartsWith("UPDATE ") Then
				Dim a As Integer = Sql2.ToUpper.IndexOf(" SET ")

				If a > 0 Then
					Dim inizio As String = Mid(Sql2, 1, a)
					Dim modificato As String = inizio.ToLower
					Sql2 = Sql2.Replace(inizio, modificato)
				End If
			Else
				Sql2 = Sql2.ToLower()
			End If
		End If

		Sql2 = Sql2.Replace("[", "")
		Sql2 = Sql2.Replace("]", "")
		Sql2 = Sql2.Replace("dbo.", "")

		Sql2 = Sql2.Replace("generale", "Generale")
		Sql2 = Sql2.Replace("dbvuoto", "dbVuoto")
		Sql2 = Sql2.Replace("dbpieno", "dbPieno")

		Return Sql2
	End Function

	'Private Function ControllaAperturaConnessione(ByRef Conn As Object, ByVal Connessione As String, Indice As Integer) As Boolean
	'	Dim Ritorno As Boolean = False

	'	If Conn Is Nothing Then
	'		If TipoDB = "SQLSERVER" Then
	'			Ritorno = True
	'			Conn = ApreDB(Connessione, Indice)
	'		Else
	'			Ritorno = True
	'			Conn = ApreDB(Connessione, Indice)
	'		End If
	'	End If

	'	Return Ritorno
	'End Function

	Public Sub ChiudeDB(Conn As Object)
		If TipoDB = "SQLSERVER" Then
			Conn.Close()
		Else
			mdb.ChiudiConn(Conn)
		End If
	End Sub

	Private Sub ThreadScriveLog(s As String, dove As String)
		Dim e As New LogStruct
		e.Cosa = s
		e.Dove = dove
		listaLog.Add(e)

		avviaTimerLog()
	End Sub

	Private Sub avviaTimerLog()
		If timerLog Is Nothing Then
			timerLog = New Timer(100)
			AddHandler timerLog.Elapsed, New ElapsedEventHandler(AddressOf scodaLog)
			timerLog.Start()
		End If
	End Sub

	Private Sub scodaLog()
		timerLog.Enabled = False
		Dim ls As LogStruct = listaLog.Item(0)
		Dim Dove As String = ls.Dove
		Dim sLog As String = ls.Cosa

		Dim gf As New GestioneFilesDirectory
		gf.CreaDirectoryDaPercorso(Dove)
		gf.ApreFileDiTestoPerScrittura(Dove)
		gf.ScriveTestoSuFileAperto(sLog)
		gf.ChiudeFileDiTestoDopoScrittura()

		listaLog.RemoveAt(0)
		If listaLog.Count > 0 Then
			timerLog.Enabled = True
		Else
			timerLog = Nothing
			listaLog = New List(Of LogStruct)
		End If
	End Sub
End Class
