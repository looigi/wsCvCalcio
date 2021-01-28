Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://cvCalcio.nfc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsNFC
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function ScriveDatiTessera(Squadra As String, NumeroTessera As String, idGiocatore As String, Descrizione As String, Importo As String) As String
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
				Dim Sql As String = "Select Max(Progressivo)+1 From TessereNFC Where NumeroTessera='" & NumeroTessera & "' And idGiocatore=" & idGiocatore
				Dim Progressivo As Integer = 0

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec(0).Value Is DBNull.Value Then
						Progressivo = 1
					Else
						Progressivo = Rec(0).Value
					End If
				End If

				Dim sImporto As String = Importo
				If sImporto = "" Then sImporto = "0"

				Dim DataOra As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

				Sql = "Insert Into TessereNFC Values (" &
					"'" & NumeroTessera & "', " &
					"'" & Squadra & "', " &
					" " & Progressivo & ", " &
					" " & idGiocatore & ", " &
					"'" & Descrizione.Replace("'", "''") & "', " &
					" " & sImporto & ", " &
					"'" & DataOra & "' " &
					")"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select Max(idLettore)+1 From LettoriNFC"
				Dim Progressivo As Integer = 0

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec(0).Value Is DBNull.Value Then
						Progressivo = 1
					Else
						Progressivo = Rec(0).Value
					End If
				End If

				Sql = "Insert Into LettoriNFC Values (" &
					" " & Progressivo & ", " &
					"'" & SistemaStringa(Descrizione) & "', " &
					"'" & SistemaStringa(IndirizzoIP) & "', " &
					"'', " &
					"'N' " &
					")"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
				If Not Ritorno.Contains(StringaErrore) Then
					Ritorno = "*"
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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Update LettoriNFC Set " &
					"Descrizione = '" & SistemaStringa(Descrizione) & "', " &
					"IndirizzoIP = '" & SistemaStringa(IndirizzoIP) & "' " &
					"Where idLettore = " & idLettore
				Ritorno = EsegueSql(Conn, Sql, Connessione)
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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Update LettoriNFC Set " &
					"Eliminato = 'S' " &
					"Where idLettore = " & idLettore
				Ritorno = EsegueSql(Conn, Sql, Connessione)
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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = "01" ' ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")

				Sql = "Select * From LettoriNFC Where Descrizione='" & NomeLettore & "'"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = "02" ' Rec
				Else
					If Rec.Eof Then
						Ritorno = "03" ' "ERROR: Lettore NFC non rilevato"
					Else
						Dim Ora As String = Now.Year & "-" & Format(Now.Month, "00") & "-" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

						Sql = "Update LettoriNFC Set " &
							"DataUltimaLettura = '" & Ora & "' " &
							"Where Descrizione = " & NomeLettore
						Ritorno = EsegueSql(Conn, Sql, Connessione)
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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim NowSpostato As DateTime = Now.AddMinutes(-30)
				Dim Ora As String = NowSpostato.Year & "-" & Format(NowSpostato.Month, "00") & "-" & Format(NowSpostato.Day, "00") & " " & Format(NowSpostato.Hour, "00") & ":" & Format(NowSpostato.Minute, "00") & ":" & Format(NowSpostato.Second, "00")
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From LettoriNFC Where Eliminato='N' Order By idLettore"

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("idLettore").Value & ";" & Rec("Descrizione").Value.replace(";", "*PV*") & ";" & Rec("IndirizzoIP").Value & ";" & Rec("DataUltimaLettura").Value & "§"

						Rec.MoveNext
					Loop
					Rec.CLose
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
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim NowSpostato As DateTime = Now.AddMinutes(-30)
				Dim Ora As String = NowSpostato.Year & "-" & Format(NowSpostato.Month, "00") & "-" & Format(NowSpostato.Day, "00") & " " & Format(NowSpostato.Hour, "00") & ":" & Format(NowSpostato.Minute, "00") & ":" & Format(NowSpostato.Second, "00")
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From LettoriNFC Where Eliminato= 'N' And Convert(DateTime, DataUltimaLettura, 121) < CONVERT(DateTime, '" & Ora & "', 121)  Order By idLettore"

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("idLettore").Value & ";" & Rec("Descrizione").Value.replace(";", "*PV*") & ";" & Rec("IndirizzoIP").Value & ";" & Rec("DataUltimaLettura").Value & "§"

						Rec.MoveNext
					Loop
					Rec.CLose
				End If

			End If
		End If

		Return Ritorno
	End Function

End Class