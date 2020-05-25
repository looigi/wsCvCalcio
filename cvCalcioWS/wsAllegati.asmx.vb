Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cv.allegati.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsAllegati
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaAllegati(Squadra As String, ByVal idAnno As String, Categoria As String, ID As String) As String
		Dim Ritorno As String = ""

		If ID = "-1" Then
			Return "Nessun allegato rilevato"
		End If

		Dim ga As New GestioneFilesDirectory
		Dim paths() As String = ga.LeggeFileIntero(Server.MapPath(".") & "\PathAllegati.txt").Split(";")
		Dim path As String = paths(0).Replace(vbCrLf, "")
		path = path.Replace(Chr(13), "")
		path = path.Replace(Chr(10), "")
		If Strings.Right(path, 1) = "\" Then
			path = Mid(path, 1, path.Length - 1)
		End If
		Dim pathLog As String = paths(1).Replace(vbCrLf, "")
		pathLog = pathLog.Replace(Chr(13), "")
		pathLog = pathLog.Replace(Chr(10), "")
		If Strings.Right(pathLog, 1) = "\" Then
			pathLog = Mid(pathLog, 1, pathLog.Length - 1)
		End If

		'Dim NomeFileLog As String = pathLog & "\LogDocumenti.txt"
		'ga.ApreFileDiTestoPerScrittura(NomeFileLog)

		path &= "\" & Squadra & "\" & Categoria & "\Anno" & idAnno & "\" & ID
		'ga.ScriveTestoSuFileAperto(Now & "->Cartella di ricerca: " & path)
		ga.CreaDirectoryDaPercorso(path & "\")
		ga.LeggeFilesDaDirectory(path)
		Dim qFiles As Integer = ga.RitornaQuantiFilesRilevati
		Dim filetti() As String = ga.RitornaFilesRilevati
		For i As Integer = 1 To qFiles
			Ritorno &= ID & ";" & filetti(i).Replace(path & "\", "") & ";§"
		Next
		If Ritorno = "" Then Ritorno = "Nessun allegato rilevato"
		'ga.ChiudeFileDiTestoDopoScrittura()

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaAllegati(Squadra As String, ByVal idAnno As String, Categoria As String, ID As String, NomeDocumento As String) As String
		Dim Ritorno As String = ""
		Dim ga As New GestioneFilesDirectory
		Dim paths() As String = ga.LeggeFileIntero(Server.MapPath(".") & "\PathAllegati.txt").Split(";")
		Dim path As String = paths(0).Replace(vbCrLf, "")
		path = path.Replace(Chr(13), "")
		path = path.Replace(Chr(10), "")
		If Strings.Right(path, 1) = "\" Then
			path = Mid(path, 1, path.Length - 1)
		End If
		Dim pathLog As String = paths(1).Replace(vbCrLf, "")
		pathLog = pathLog.Replace(Chr(13), "")
		pathLog = pathLog.Replace(Chr(10), "")
		If Strings.Right(pathLog, 1) = "\" Then
			pathLog = Mid(pathLog, 1, pathLog.Length - 1)
		End If
		Dim NomeFileLog As String = pathLog & "\LogDocumenti.txt"
		ga.ApreFileDiTestoPerScrittura(NomeFileLog)

		path &= "\" & Squadra & "\" & Categoria & "\Anno" & idAnno & "\" & ID & "\" & NomeDocumento
		ga.ScriveTestoSuFileAperto(Now & "->Eliminazione allegato: " & path)

		Try
			Kill(path)
			ga.ScriveTestoSuFileAperto(Now & "->Eseguito")
			Ritorno = "*"
		Catch ex As Exception
			ga.ScriveTestoSuFileAperto(Now & "->" & StringaErrore & ex.Message)
			Ritorno = StringaErrore & ex.Message
		End Try

		ga.ChiudeFileDiTestoDopoScrittura()
		ga = Nothing

		Return Ritorno
	End Function
End Class