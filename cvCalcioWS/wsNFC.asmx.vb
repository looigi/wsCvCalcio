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

End Class