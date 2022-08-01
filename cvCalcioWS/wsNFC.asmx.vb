Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://cvCalcio.nfc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsNFC
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function ValidaCashback(NumeroTessera As String, Importo As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Max(Progressivo),0)+1", "Coalesce(Max(Progressivo),0)+1") & " From cashbackutilizzato Where codicetessera='" & NumeroTessera & "'"
				Dim Progressivo As Integer = 0

				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Progressivo = Rec(0).Value

					Rec.Close

					Dim DataOra As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

					Sql = "Insert Into cashbackutilizzato Values ('" & NumeroTessera & "', " & Progressivo & ", " & Importo & ", '" & DataOra & "')"

					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Not Ritorno.Contains(StringaErrore) Then
						Ritorno = "*"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ScriveDatiTessera(NumeroTessera As String, Descrizione As String, Importo As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Max(Progressivo),0)+1", "Coalesce(Max(Progressivo),0)+1") & " From TessereNFC Where NumeroTessera='" & NumeroTessera & "'" '  And idGiocatore=" & idGiocatore
				Dim Progressivo As Integer = 0

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	Progressivo = 1
					'Else
					Progressivo = Rec(0).Value
					'End If
				End If

				Dim sImporto As String = Importo
				If sImporto = "" Then sImporto = "0"

				Dim DataOra As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

				'"'" & Squadra & "', " &
				'" " & idGiocatore & ", " &

				Sql = "Insert Into TessereNFC Values (" &
					"'" & NumeroTessera & "', " &
					" " & Progressivo & ", " &
					"'" & Descrizione.Replace("'", "''") & "', " &
					" " & sImporto & ", " &
					"'" & DataOra & "' " &
					")"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Not Ritorno.Contains(StringaErrore) Then
					Ritorno = "*"
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function NuovoLettoreNFC(Squadra As String, Descrizione As String, IndirizzoIP As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Max(idLettore),0)+1", "Coalesce(Max(idLettore),0)+1") & " From LettoriNFC"
				Dim Progressivo As Integer = 0

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	Progressivo = 1
					'Else
					Progressivo = Rec(0).Value
					'End If
				End If

				Sql = "Insert Into LettoriNFC Values (" &
					" " & Progressivo & ", " &
					"'" & SistemaStringa(Descrizione) & "', " &
					"'" & SistemaStringa(IndirizzoIP) & "', " &
					"'', " &
					"'N' " &
					")"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Not Ritorno.Contains(StringaErrore) Then
					Ritorno = "*"
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaDatiDaTessera(CodiceTessera As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB("Generale")

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From GiocatoriTessereNFC Where CodiceTessera='" & CodiceTessera & "'"
				Dim Progressivo As Integer = 0

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof = False Then
						Ritorno = Rec("CodSquadra").Value & ";" & Rec("idGiocatore").Value & ";"
					Else
						Ritorno = "ERROR: Nessuna tessera rilevata con il codice " & CodiceTessera
					End If
					Rec.Close()
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaLettoreNFC(Squadra As String, idLettore As String, Descrizione As String, IndirizzoIP As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Update LettoriNFC Set " &
					"Descrizione = '" & SistemaStringa(Descrizione) & "', " &
					"IndirizzoIP = '" & SistemaStringa(IndirizzoIP) & "' " &
					"Where idLettore = " & idLettore
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Not Ritorno.Contains(StringaErrore) Then
					Ritorno = "*"
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaLettoreNFC(Squadra As String, idLettore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Update LettoriNFC Set " &
					"Eliminato = 'S' " &
					"Where idLettore = " & idLettore
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Not Ritorno.Contains(StringaErrore) Then
					Ritorno = "*"
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function WatchdogLettoreNFC(Squadra As String, NomeLettore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = "01" ' ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = "01" ' ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object

				Sql = "Select * From LettoriNFC Where Descrizione='" & NomeLettore & "'"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = "02" ' Rec
				Else
					If Rec.Eof() Then
						Ritorno = "03" ' "ERROR: Lettore NFC non rilevato"
					Else
						Dim Ora As String = Now.Year & "-" & Format(Now.Month, "00") & "-" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

						Sql = "Update LettoriNFC Set " &
							"DataUltimaLettura = '" & Ora & "' " &
							"Where Descrizione = " & NomeLettore
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Not Ritorno.Contains(StringaErrore) Then
							Ritorno = "*"
						Else
							Ritorno = "13" ' Errore nell'update
						End If
					End If
				End If
			End If
		End If

		If Ritorno = "*" Then
			Ritorno = "00"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaLettoriNFC(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim NowSpostato As DateTime = Now.AddMinutes(-30)
				Dim Ora As String = NowSpostato.Year & "-" & Format(NowSpostato.Month, "00") & "-" & Format(NowSpostato.Day, "00") & " " & Format(NowSpostato.Hour, "00") & ":" & Format(NowSpostato.Minute, "00") & ":" & Format(NowSpostato.Second, "00")
				Dim Rec As Object
				Dim Sql As String = "Select * From LettoriNFC Where Eliminato='N' Order By idLettore"

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof()
						Ritorno &= Rec("idLettore").Value & ";" & Rec("Descrizione").Value.replace(";", "*PV*") & ";" & Rec("IndirizzoIP").Value & ";" & Rec("DataUltimaLettura").Value & "§"

						Rec.MoveNext()
					Loop
					Rec.Close()
				End If

			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaLettoreNFCOffLine(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim NowSpostato As DateTime = Now.AddMinutes(-30)
				Dim Ora As String = NowSpostato.Year & "-" & Format(NowSpostato.Month, "00") & "-" & Format(NowSpostato.Day, "00") & " " & Format(NowSpostato.Hour, "00") & ":" & Format(NowSpostato.Minute, "00") & ":" & Format(NowSpostato.Second, "00")
				Dim Rec as Object
				Dim Sql As String = "Select * From LettoriNFC Where Eliminato= 'N' And " &
					" " & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, DataUltimaLettura, 121)", "Convert(DataUltimaLettura, DateTime)") & " < " & IIf(TipoDB = "SQLSERVER", "CONVERT(DateTime, '" & Ora & "', 121)", "CONVERT('" & Ora & "', DateTime)") & " Order By idLettore"

				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof()
						Ritorno &= Rec("idLettore").Value & ";" & Rec("Descrizione").Value.replace(";", "*PV*") & ";" & Rec("IndirizzoIP").Value & ";" & Rec("DataUltimaLettura").Value & "§"

						Rec.MoveNext()
					Loop
					Rec.Close()
				End If

			End If
		End If

		Return Ritorno
	End Function

End Class