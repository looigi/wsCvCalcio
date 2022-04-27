Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_mult.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsMultimedia
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function EliminaImmagine(Squadra As String, Tipologia As String, NomeFile As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Righe As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim Campi() As String = Righe.Split(";")

		Dim Ritorno As String = ""
		Dim Ok As Boolean = True

		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
		filePaths = filePaths.Replace(vbCrLf, "")
		If Strings.Right(filePaths, 1) <> "\" Then
			filePaths &= "\"
		End If
		Dim Percorso As String = filePaths & Squadra & "\" & Tipologia

		Dim Estensioni() As String = {".kgb", ".png", ".bmp", ".jpeg", ".jpg"}
		Dim Estensione As String
		Dim Nome As String

		If NomeFile.Contains(".") Then
			Estensione = Mid(NomeFile, NomeFile.IndexOf(".") + 1, NomeFile.Length)
			Nome = NomeFile.Replace(Estensione, "")
		Else
			Estensione = ""
			Nome = NomeFile
		End If

		For Each est As String In Estensioni
			Dim Nometto As String = Percorso & "\" & Nome & est
			If ControllaEsistenzaFile(Nometto) Then
				Try
					File.Delete(Nometto)
					Ritorno = "*"
					Ok = True
				Catch ex As Exception
					Ok = False
					Ritorno = StringaErrore & " " & ex.Message
				End Try
				Exit For
			End If
		Next

		If Ritorno = "" Then
			Ritorno = "*"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMultimedia(Squadra As String, idAnno As String, id As String, Tipologia As String) As String
		Return RitornaMultimediaPerTipologia(Server.MapPath("."), Squadra, idAnno, id, Tipologia)
	End Function

	<WebMethod()>
	Public Function EliminaMultimedia(Immagine As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim paths() As String = gf.LeggeFileIntero(Server.MapPath(".") & "/Impostazioni/PathAllegati.txt").Split(";")
		Dim PathIniziale As String = paths(0).Replace(vbCrLf, "")
		If Strings.Right(PathIniziale, 1) <> "\" Then
			PathIniziale &= "\"
		End If
		Dim Ritorno As String = ""

		If ControllaEsistenzaFile(PathIniziale & Immagine) Then
			Ritorno = gf.EliminaFileFisico(PathIniziale & Immagine)
		End If

		gf = Nothing

		If Ritorno = "" Then Ritorno = "*"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaAlbumPerCategoria(Squadra As String, idAnno As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Ok As Boolean = True
		Dim gf As New GestioneFilesDirectory
		Dim PathIniziale As String = gf.LeggeFileIntero(Server.MapPath(".") & "/Impostazioni/Paths.txt")
		PathIniziale = PathIniziale.Trim
		If Not PathIniziale.EndsWith("\") Then
			PathIniziale &= "\"
		End If

		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Altro As String = ""
				Dim Rec as object
				Dim Sql As String = ""

				If idCategoria <> "-1" Then
					Altro = " And Partite.idCategoria = " & idCategoria
				End If

				Try
					Sql = "Select Partite.idPartita, Partite.DataOra, Partite.Casa, SquadreAvversarie.Descrizione, Sum(Goal) As Segnati, Sum(GoalAvversari) As Subiti  From ( "
					Sql &= "Select Partite.idPartita, Partite.DataOra, Partite.Casa, SquadreAvversarie.Descrizione, 0 As Goal, Sum(IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) As GoalAvversari "
					Sql &= "From (Partite Left Join SquadreAvversarie On Partite.idAvversario=SquadreAvversarie.idAvversario) "
					Sql &= "Left Join RisultatiAggiuntivi On Partite.idPartita=RisultatiAggiuntivi.idPartita "
					Sql &= "Where Partite.idAnno=" & idAnno & Altro & " "
					Sql &= "Group By Partite.idPartita, Partite.DataOra, Partite.Casa, SquadreAvversarie.Descrizione "
					Sql &= "Union All "
					Sql &= "Select Partite.idPartita, Partite.DataOra, Partite.Casa, SquadreAvversarie.Descrizione, " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " As Goal, 0 As GoalAvversari "
					Sql &= "From (Partite Left Join SquadreAvversarie On Partite.idAvversario=SquadreAvversarie.idAvversario) "
					Sql &= "Left Join RisultatiAggiuntiviMarcatori On Partite.idPartita=RisultatiAggiuntiviMarcatori.idPartita "
					Sql &= "Where Partite.idAnno=" & idAnno & Altro & " "
					Sql &= "Group By Partite.idPartita, Partite.DataOra, Partite.Casa, SquadreAvversarie.Descrizione "
					Sql &= ") As A Group By Partite.idPartita, Partite.DataOra, Partite.Casa, SquadreAvversarie.Descrizione"

					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						'Ritorno = Rec
					Else
						If Not Rec.Eof() Then
							Dim idPartite As List(Of Integer) = New List(Of Integer)
							Dim Desc As List(Of String) = New List(Of String)

							Do Until Rec.Eof()
								Dim Casa As String = ""
								Select Case Rec("Casa").Value
									Case "S"
										Casa = "In casa"
									Case "N"
										Casa = "Fuori casa"
									Case "E"
										Casa = "Campo esterno"
								End Select
								idPartite.Add(Rec("idPartita").Value)
								Desc.Add("Partite;" & Rec("DataOra").Value & ";" & Casa & ";" & Rec("Descrizione").Value & ";" & Rec("Segnati").Value & "-" & Rec("Subiti").Value & ";")

								Rec.MoveNext()
							Loop
							Rec.Close()

							Dim Partita As Integer = 0

							' Ritorno = ""

							For Each i As Integer In idPartite
								Dim Path As String = PathIniziale & "Partite\" & i
								gf.ScansionaDirectorySingola(Path)
								Dim Filetti() As String = gf.RitornaFilesRilevati
								Dim qFiletti As String = gf.RitornaQuantiFilesRilevati

								' Ritorno &= PathIniziale & "Partite\" & i

								For k As Integer = 1 To qFiletti
									Ritorno &= Filetti(k).Replace(PathIniziale, "") & ";" & Desc.Item(Partita) & "§"
								Next

								Partita += 1
							Next
						End If

						' Return Ritorno


						If idCategoria <> "-1" Then
							Altro = " And idCategoria = " & idCategoria
						End If

						Sql = "Select Giocatori.*, Ruoli.Descrizione From Giocatori "
						Sql &= "Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo "
						Sql &= "Where idAnno=" & idAnno & Altro
						Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
						If TypeOf (Rec) Is String Then
							'Ritorno = Rec
						Else
							If Not Rec.Eof() Then
								Dim idGiocatori As List(Of Integer) = New List(Of Integer)
								Dim Desc As List(Of String) = New List(Of String)

								Do Until Rec.Eof()
									idGiocatori.Add(Rec("idGiocatore").Value)
									Desc.Add("Giocatori;" & Rec("Cognome").Value & " " & Rec("Nome").Value & ";" & Rec("Soprannome").Value & ";" & Rec("Descrizione").Value & ";" & Rec("DataDiNascita").Value & ";")

									Rec.MoveNext()
								Loop
								Rec.Close()

								Dim Giocatore As Integer = 0

								For Each i As Integer In idGiocatori
									Dim Path As String = PathIniziale & "Giocatori\" & idAnno & "_" & i.ToString
									gf.ScansionaDirectorySingola(Path)
									Dim Filetti() As String = gf.RitornaFilesRilevati
									Dim qFiletti As String = gf.RitornaQuantiFilesRilevati
									For k As Integer = 1 To qFiletti
										Ritorno &= Filetti(k).Replace(PathIniziale, "") & ";" & Desc.Item(Giocatore) & "§"
									Next
									Giocatore += 1
								Next
							End If
						End If

						Sql = "Select * From Allenatori Where idAnno=" & idAnno & Altro
						Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
						If TypeOf (Rec) Is String Then
							'Ritorno = Rec
						Else
							If Not Rec.Eof() Then
								Dim idAllenatore As List(Of Integer) = New List(Of Integer)
								Dim Desc As List(Of String) = New List(Of String)

								Do Until Rec.Eof()
									idAllenatore.Add(Rec("idAllenatore").Value)
									Desc.Add("Allenatori;" & Rec("Cognome").Value & " " & Rec("Nome").Value & ";;;;")

									Rec.MoveNext()
								Loop
								Rec.Close()

								Dim Allenatore As Integer = 0

								For Each i As Integer In idAllenatore
									Dim Path As String = PathIniziale & "Allenatori\" & idAnno & "_" & i.ToString & ".jpg"
									'gf.ScansionaDirectorySingola(Path)
									'Dim Filetti() As String = gf.RitornaFilesRilevati
									'Dim qFiletti As String = gf.RitornaQuantiFilesRilevati
									'For k As Integer = 1 To qFiletti
									If ControllaEsistenzaFile(Path) Then
										Ritorno &= Path.Replace(PathIniziale, "") & ";" & Desc.Item(Allenatore) & "§"
									End If
									'Next
									Allenatore += 1
								Next
							End If
						End If

						Sql = "Select * From Categorie Where idAnno=" & idAnno & Altro
						Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
						If TypeOf (Rec) Is String Then
							'Ritorno = Rec
						Else
							If Not Rec.Eof() Then
								Dim idCategoria2 As List(Of Integer) = New List(Of Integer)
								Dim Desc As List(Of String) = New List(Of String)

								Do Until Rec.Eof()
									idCategoria2.Add(Rec("idCategoria").Value)
									Desc.Add("Categorie;" & Rec("Descrizione").Value & ";;;;")

									Rec.MoveNext()
								Loop
								Rec.Close()

								Dim Categoria As Integer = 0

								For Each i As Integer In idCategoria2
									Dim Path As String = PathIniziale & "Categorie\" & idAnno & "_" & i.ToString & ".jpg"
									'    gf.ScansionaDirectorySingola(Path)
									'    Dim Filetti() As String = gf.RitornaFilesRilevati
									'    Dim qFiletti As String = gf.RitornaQuantiFilesRilevati
									'    For k As Integer = 1 To qFiletti
									If ControllaEsistenzaFile(Path) Then
										Ritorno &= Path.Replace(PathIniziale, "") & ";" & Desc.Item(Categoria) & "§"
									End If
									'    Next
									Categoria += 1
								Next
							End If
						End If

						'Sql = "Select * From Dirigenti Where idAnno=" & idAnno & Altro
						'Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
						'If TypeOf (Rec) Is String Then
						'    'Ritorno = Rec
						'Else
						'    If Not Rec.Eof() Then
						'        Dim idDirigente As List(Of Integer) = New List(Of Integer)
						'        Dim Desc As List(Of String) = New List(Of String)

						'        Do Until Rec.Eof()
						'            idDirigente.Add(Rec("idDirigente").Value)
						'            Desc.Add("Dirigenti;" & Rec("Cognome").Value & " " & Rec("Nome").Value & ";;;;")

						'            Rec.MoveNext()
						'        Loop
						'        Rec.Close()

						'        Dim Dirigente As Integer = 0

						'        For Each i As Integer In idDirigente
						'            Dim Path As String = PathIniziale & "Dirigenti\" & i.ToString
						'            gf.ScansionaDirectorySingola(Path)
						'            Dim Filetti() As String = gf.RitornaFilesRilevati
						'            Dim qFiletti As String = gf.RitornaQuantiFilesRilevati
						'            For k As Integer = 1 To qFiletti
						'                Ritorno &= Filetti(k).Replace(PathIniziale, "") & ";" & Desc.Item(Dirigente) & "§"
						'            Next
						'            Dirigente += 1
						'        Next
						'    End If
						'End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		If Ritorno = "" Then
			Ritorno = StringaErrore & " Nessun multimedia rilevato"
		End If

		Return Ritorno
	End Function

End Class