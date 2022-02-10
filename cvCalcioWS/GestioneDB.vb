Imports System.Reflection
Imports System.Timers
Imports ADODB
Imports MySqlConnector

Public Class clsGestioneDB
	Private mdb As clsMariaDB

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

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			'If nomeFileLogGenerale = "" Then
			Dim paths As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")
			Dim pp() As String = paths.Split(";")
			pp(1) = pp(1).Replace(vbCrLf, "")
			If Strings.Right(pp(1), 1) <> "\" Then
				pp(1) = pp(1) & "\"
			End If
			nomeFileLogGenerale = pp(1) & "logWS_Exec_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
			ThreadScriveLog(Datella & "--------------------------------------------------------------------------")

			Dim Sql2 As String = ""

			If ModificaQuery Then
				If TipoDB = "SQLSERVER" Then
					Sql2 = Sql
				Else
					Sql2 = Sql.ToLower()
					Sql2 = Sql2.Replace("[", "")
					Sql2 = Sql2.Replace("]", "")
					Sql2 = Sql2.Replace("dbo.", "")

					Sql2 = Sql2.Replace("generale", "Generale")
				End If
			Else
				Sql2 = Sql
			End If

			ThreadScriveLog(Datella & ": " & Sql2)
			' End If
		End If

		' Routine che esegue una query sul db
		If TipoDB = "SQLSERVER" Then
			Try
				Conn.Execute(Sql)
				If effettuaLog Then
					ThreadScriveLog(Datella & ": OK")
				End If
			Catch ex As Exception
				Ritorno = StringaErrore & " " & ex.Message
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message)
				End If

			End Try
		Else
			Try
				Ritorno = mdb.EsegueSql(Sql, ModificaQuery)
				If Ritorno.ToUpper <> "OK" Then
					Ritorno = StringaErrore & " " & Ritorno
				End If
				If effettuaLog Then
					ThreadScriveLog(Datella & ": OK")
				End If
			Catch ex As Exception
				Ritorno = StringaErrore & " " & ex.Message
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message)
				End If

			End Try
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			ThreadScriveLog(Datella & "--------------------------------------------------------------------------")
			ThreadScriveLog("")
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

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			'If nomeFileLogGenerale = "" Then
			Dim paths As String = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")

			Dim pp() As String = paths.Split(";")
			pp(1) = pp(1).Replace(vbCrLf, "")
			If Strings.Right(pp(1), 1) <> "\" Then
				pp(1) = pp(1) & "\"
			End If
			nomeFileLogGenerale = pp(1) & "logWS_Query_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"

			ThreadScriveLog(Datella & "--------------------------------------------------------------------------")
			ThreadScriveLog(Datella & " TIPO DB: " & TipoDB)

			Dim Sql2 As String = ""

			If ModificaQuery Then
				If TipoDB = "SQLSERVER" Then
					Sql2 = Sql
				Else
					Sql2 = Sql.ToLower()
					Sql2 = Sql2.Replace("[", "")
					Sql2 = Sql2.Replace("]", "")
					Sql2 = Sql2.Replace("dbo.", "")

					Sql2 = Sql2.Replace("generale", "Generale")
				End If
			Else
				Sql2 = Sql
			End If

			ThreadScriveLog(Datella & ": " & Sql2)
			'End If
		End If

		'Return "Lettura " & Indice & " -> " & mdb.Length

		Dim Rec As Object

		If TipoDB = "SQLSERVER" Then
			Rec = New Recordset

			Try
				Rec.Open(Sql, Conn)
			Catch ex As Exception
				Rec = StringaErrore & " " & ex.Message
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message)
				End If
			End Try
		Else
			Try
				Rec = mdb.Lettura(Sql, ModificaQuery)
			Catch ex As Exception
				If effettuaLog Then
					ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message)
				End If
				Return StringaErrore & " " & ex.Message
			End Try
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			ThreadScriveLog(Datella & "--------------------------------------------------------------------------")
			ThreadScriveLog("")
		End If

		ChiudeDB(Conn)

		Return Rec
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

	Private Sub ThreadScriveLog(s As String)
		listaLog.Add(s)

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
		Dim sLog As String = listaLog.Item(0)

		Dim gf As New GestioneFilesDirectory
		gf.ApreFileDiTestoPerScrittura(nomeFileLogGenerale)
		gf.ScriveTestoSuFileAperto(sLog)
		gf.ChiudeFileDiTestoDopoScrittura()

		listaLog.RemoveAt(0)
		If listaLog.Count > 0 Then
			timerLog.Enabled = True
		Else
			timerLog = Nothing
			listaLog = New List(Of String)
		End If
	End Sub
End Class
