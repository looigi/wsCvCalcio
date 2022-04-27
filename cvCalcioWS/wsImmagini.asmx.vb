Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB
Imports System.Windows.Forms
Imports System.IO

<System.Web.Services.WebService(Namespace:="http://cvcalcio_imm.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsImmagini
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function ConverteImmaginiVersoDB(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim ConnessioneGenerale As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim ConnGen As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim NomeSquadra As String = ""
				Dim Ok As Boolean = True

				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From Squadre Where idSquadra = " & Val(ss(1)).ToString
				Rec = ConnGen.LeggeQuery(Server.MapPath("."), Sql, ConnessioneGenerale)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Squadra non rilevata: " & Squadra
						Ok = False
					Else
						NomeSquadra = "" & Rec("Descrizione").Value
					End If
				End If
				Rec.Close()

				If Ok Then
					'Dim Categorie() As String = {"Allenatori", "Arbitri", "Avversari", "Categorie", "Dirigenti", "Firme", "Giocatori", "Societa"}
					Dim gf As New GestioneFilesDirectory
					'Dim PathSquadra As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					'PathSquadra = PathSquadra.Replace(vbCrLf, "")
					'PathSquadra &= "/" & NomeSquadra.Replace(" ", "_") & "/"

					Dim Filetti() As String
					Dim qFiles As Long

					'For Each c As String In Categorie
					'	Dim Path As String = PathSquadra & c
					'	gf.ScansionaDirectorySingola(Path)
					'	Filetti = gf.RitornaFilesRilevati
					'	qFiles = gf.RitornaQuantiFilesRilevati
					'	For i As Long = 1 To qFiles
					'		Dim Immagine As String = DecriptaImmagine(Server.MapPath("."), Filetti(i))
					'		Immagine = Immagine.Replace("http://192.168.0.205:1011/", "/var/www/html/inCalcio/Multimedia/")
					'		Dim NomeFile As String = gf.TornaNomeFileDaPath(Filetti(i))

					'		Ritorno = SalvaImmagineDB(Squadra, c, Immagine, NomeFile, "NO")
					'		' Ritorno &= Squadra & ";" & c & ";" & Immagine & ";" & NomeFile & "§"
					'	Next
					'Next

					' ALLEGATI
					'Dim Partite As String = "Partite"
					'Dim p As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
					'p = p.Replace(vbCrLf, "")
					'Dim pp() As String = p.Split(";")
					'Dim PathSquadraAllegati As String = pp(0) & "/" & Squadra & "/" & Partite & "/"

					'gf.ScansionaDirectorySingola(PathSquadraAllegati)
					'Filetti = gf.RitornaFilesRilevati
					'qFiles = gf.RitornaQuantiFilesRilevati
					'For i As Long = 1 To qFiles
					'	If Filetti(i).ToUpper.Contains(".PNG") Or Filetti(i).ToUpper.Contains(".JPG") Or Filetti(i).ToUpper.Contains(".JPEG") Then
					'		Dim a As Integer = Filetti(i).IndexOf("Anno1/")
					'		Dim CodicePartita As String = Mid(Filetti(i), a + 7, Filetti.Length)
					'		a = CodicePartita.IndexOf("/")
					'		CodicePartita = Mid(CodicePartita, 1, a)

					'		'Dim Immagine As String = DecriptaImmagine(Server.MapPath("."), Filetti(i))
					'		'Immagine = Immagine.Replace("http://192.168.0.205:1011/", "/var/www/html/inCalcio/Multimedia/")
					'		Dim NomeFile As String = gf.TornaNomeFileDaPath(Filetti(i))
					'		CodicePartitaPerImmagini = CodicePartita

					'		' Ritorno = SalvaImmagineDB(Squadra, "Partite", Filetti(i), NomeFile, "SI")
					'		Ritorno &= Filetti(i) & "->" & CodicePartitaPerImmagini & ";"
					'	End If
					'Next i
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaAllegati(Squadra As String, Tipologia As String, ID As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim ConnGen As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim NomePrefisso As String = ""

				If Tipologia = "Partite" Then
					NomePrefisso = "immagini"
				Else
					NomePrefisso = "allegati"
				End If

				Dim Rec As Object
				Dim Sql As String = "Select * From " & NomePrefisso & "_" & Tipologia & " Where Id=" & ID & " Order By Progressivo"

				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Ritorno = ""
					Do Until Rec.Eof()
						If Tipologia = "Partite" Then
							Ritorno &= Rec("Dati").Value & "^" & Rec("NomeFile").Value & "^" & Rec("Lunghezza").Value & "^^" & Rec("Progressivo").Value & "^§"
						Else
							Ritorno &= "" & ";" & Rec("NomeFile").Value & ";" & Rec("Lunghezza").Value & ";;" & Rec("Progressivo").Value & ";§"
						End If

						Rec.MoveNext()
					Loop
					Rec.Close()
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaAllegato(Squadra As String, Tipologia As String, ID As String, Progressivo As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim ConnGen As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim NomePrefisso As String = ""

				If Tipologia = "Partite" Then
					NomePrefisso = "immagini"
				Else
					NomePrefisso = "allegati"
				End If

				Dim Sql As String = "Delete From " & NomePrefisso & "_" & Tipologia & " Where Id=" & ID & " And Progressivo=" & Progressivo

				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				If Not Ritorno.Contains(StringaErrore) Then
					Ritorno = "*"
				End If
			End If
		End If

		Return Ritorno
	End Function

	Private Function UnicodeBytesToString(ByVal bytes() As Byte) As String

		Return System.Text.Encoding.UTF8.GetString(bytes)
	End Function

	<WebMethod()>
	Public Function SalvaImmagineDB(Squadra As String, Tipologia As String, PathImmagine As String, NomeFile As String, Allegati As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim ConnGen As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim gf As New GestioneFilesDirectory
				Dim Este As String = gf.TornaEstensioneFileDaPath(NomeFile).ToUpper.Trim
				Dim id As String
				Dim Rec As Object
				Dim b64 As String = ""

				If Allegati = "SI" Then
					If PathImmagine.Contains("_") Then
						Dim Splittone() As String = PathImmagine.Split("_")
						id = Val(Splittone(1))
					Else
						'If CodicePartitaPerImmagini <> "" Then
						'	id = CodicePartitaPerImmagini
						'Else
						Return StringaErrore & " Problemi nell'ottenere il codice del multimedia"
						'End If
					End If

					If id.Contains(".") Then
						id = Mid(id, 1, id.IndexOf("."))
					End If

					If Este.Contains("JPG") Or Este.Contains("GIF") Or Este.Contains("JPEG") Or Este.Contains("BMP") Or Este.Contains("PNG") Then
						b64 = ConverteImmagineBase64(PathImmagine)
					Else
						b64 = UnicodeBytesToString(File.ReadAllBytes(PathImmagine)).Replace("'", "''")
					End If

					Dim NomePrefisso As String = ""

					If Tipologia = "Partite" Then
						NomePrefisso = "immagini"
						NomeFile = NomeFile.Replace(Este, "")
					Else
						NomePrefisso = "allegati"
					End If

					If TipoDB = "SQLSERVER" Then
						Sql = "Select IsNull(max(Progressivo),0)+1 From " & NomePrefisso & "_" & Tipologia.ToLower & " Where id=" & id
					Else
						Sql = "Select Coalesce(max(Progressivo),0)+1 From " & NomePrefisso & "_" & Tipologia.ToLower & " Where id=" & id
					End If
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Dim Progressivo As Integer = Rec(0).Value
						Rec.Close()

						Sql = "INSERT INTO " & NomePrefisso & "_" & Tipologia.ToLower & " (id, Progressivo, Lunghezza, Dati, NomeFile) VALUES(" & id & ", " & Progressivo & ", " & b64.Length & ", '" & b64 & "', '" & NomeFile.Replace("'", "''") & "');"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Not Ritorno.Contains(StringaErrore) Then
							gf.EliminaFileFisico(PathImmagine)
							Ritorno = "*"
						End If
					End If
				Else
					Dim Privacy As String = "n"
					Dim Progressivo As Integer = -1

					NomeFile = NomeFile.Replace(Este, "")
					If NomeFile.Contains("_") Then
						Dim Splittone() As String = NomeFile.Split("_")
						id = Splittone(1)
						If Tipologia = "Firme" Then
							Progressivo = Val(Splittone(2).Replace(Este, ""))
						End If
					Else
						id = NomeFile
					End If
					b64 = ConverteImmagineBase64(PathImmagine)
					'Return b64

					If id.Contains(".") Then
						id = Mid(id, 1, id.IndexOf("."))
					End If

					If Tipologia = "Firme" Then
						If NomeFile.Contains("_P") Then
							Privacy = "s"
						End If

						Sql = "Select * From immagini_" & Tipologia.ToLower & " Where id=" & id & " And Progressivo=" & Progressivo & " And Privacy='" & Privacy & "'"
					Else
						Sql = "Select * From immagini_" & Tipologia.ToLower & " Where id=" & id
					End If
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Tipologia = "Firme" Then
							If Not Rec.Eof Then
								Sql = "Update immagini_" & Tipologia.ToLower & " Set Lunghezza=" & b64.Length & ", Dati='" & b64 & "', Privacy='" & Privacy & "' Where id=" & id & " And Progressivo=" & Progressivo & " And Privacy='" & Privacy & "';"
							Else
								Sql = "INSERT INTO immagini_" & Tipologia.ToLower & " (id, Progressivo, Lunghezza, Dati, Privacy) VALUES(" & id & ", " & Progressivo & ", " & b64.Length & ", '" & b64 & "', '" & Privacy & "');"
							End If
						Else
							If Not Rec.Eof Then
								Sql = "Update immagini_" & Tipologia.ToLower & " Set Lunghezza=" & b64.Length & ", Dati='" & b64 & "' Where id=" & id & ";"
							Else
								Sql = "INSERT INTO immagini_" & Tipologia.ToLower & " (id, Lunghezza, Dati) VALUES(" & id & ", " & b64.Length & ", '" & b64 & "');"
							End If
						End If
						Rec.Close()
					End If

					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Not Ritorno.Contains(StringaErrore) Then
						gf.EliminaFileFisico(PathImmagine)
						Ritorno = "*"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Private Function ConverteImmagineBase64(path As String) As String
		Using image As System.Drawing.Image = System.Drawing.Image.FromFile(path)
			Using m As MemoryStream = New MemoryStream()
				image.Save(m, image.RawFormat)
				Dim imageBytes As Byte() = m.ToArray()
				Dim base64String As String = Convert.ToBase64String(imageBytes)

				Return base64String
			End Using
		End Using
	End Function

	<WebMethod()>
	Public Function RitornaImmagineDB(Squadra As String, Tipologia As String, Id As String) As String
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
				Dim Rec As Object

				Sql = "Select * From immagini_" & Tipologia.ToLower & " Where id=" & Id
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Ritorno = Rec("Dati").Value
					Else
						Ritorno = StringaErrore & " Nessuna immagine rilevata"
					End If
					Rec.Close()
				End If

			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaImmagineDB(Squadra As String, Tipologia As String, Id As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim ConnGen As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object

				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Sql = "Delete From immagini_" & Tipologia.ToLower & " Where id=" & Id
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Not Ritorno.Contains(StringaErrore) Then
						Ritorno = "*"
					End If
				End If

			End If
		End If

		Return Ritorno
	End Function
End Class