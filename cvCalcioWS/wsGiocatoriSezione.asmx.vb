Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://giocsez.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsGiocatoriSezione
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaDatiGiocatore(Squadra As String, idUtente As String) As String
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
				Dim idGiocatore As String = ""
				Dim Ok As Boolean = True

				Dim gf As New GestioneFilesDirectory
				Dim path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim pp() As String = path.Split(";")
				If Strings.Right(pp(0), 1) <> "\" Then
					pp(0) = pp(0) & "\"
				End If
				Dim a() As String = Squadra.Split("_")
				Dim Anno As Integer = Val(a(0))

				Dim p As String = pp(0) & Squadra & "\Certificati\Anno" & Anno & "\" & idUtente & "\"
				gf.ScansionaDirectorySingola(p)
				Dim filetti() As String = gf.RitornaFilesRilevati
				Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

				Ritorno &= idUtente & ";"
				For i As Integer = 1 To qFiletti
					Ritorno &= gf.TornaNomeFileDaPath(filetti(i)) & ";"
				Next
				Ritorno &= "§"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaRicevuteGiocatore(Squadra As String, idUtente As String) As String
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
				Dim idGiocatore As String = ""
				Dim Ok As Boolean = True

				Dim gf As New GestioneFilesDirectory
				Dim path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim pp() As String = path.Split(";")
				If Strings.Right(pp(0), 1) <> "\" Then
					pp(0) = pp(0) & "\"
				End If
				Dim a() As String = Squadra.Split("_")
				Dim Anno As Integer = Val(a(0))

				Dim p As String = pp(0) & Squadra & "\Ricevute\Anno" & Anno & "\" & idUtente & "\"
				gf.ScansionaDirectorySingola(p)
				Dim filetti() As String = gf.RitornaFilesRilevati
				Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati

				Ritorno = ""
				For i As Integer = 1 To qFiletti
					If filetti(i).ToUpper.Contains(".PDF") Then
						Ritorno &= idUtente & ";"
						Ritorno &= gf.TornaNomeFileDaPath(filetti(i)) & ";"
						Ritorno &= "§"
					End If
				Next
			End If
		End If

		Return Ritorno
	End Function
End Class