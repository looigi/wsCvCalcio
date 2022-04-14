Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://wsPermessiUtente.PAndE.it/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsPermessiUtente
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function SalvaPermessiUtente(Squadra As String, IDutente As Integer, Permessi As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Ok As Boolean = True
		Dim Sql As String = ""

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Delete From PermessiUtente Where Idutente = " & IDutente
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						If Permessi.Length > 0 Then
							Dim Perm() As String = Permessi.Split(",")
							Dim Progressivo As Integer = 0

							For Each p As String In Perm
								If p <> "" Then
									Progressivo += 1

									Try
										Sql = "Insert Into PermessiUtente Values (" &
											" " & IDutente & ", " &
											" " & Progressivo & ", " &
											" " & p & " " &
											")"
										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
											Exit For
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & ex.Message
										Ok = False
										Exit For
									End Try
								End If
							Next

						End If
					End If
				End If

			End If

			If Ok Then
				Sql = "commit"
				Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			Else
				Sql = "rollback"
				Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaPermessiUtente(Squadra As String, IDutente As Integer, idTipologia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					Sql = "SELECT A.*, C.Descrizione, C.NomePerCodice From PermessiUtente A " &
						"Left Join [Generale].[dbo].[Permessi_Composizione] B On A.idPermesso = B.idPermesso " &
						"Left Join [Generale].[dbo].[Permessi_Lista] C On B.idPermesso = C.idPermesso " &
						"Where A.IDutente=" & IDutente & " And B.idTipologia=" & idTipologia & " " &
						"Order By progressivo"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = "" ' StringaErrore & " Nessun permesso ritornato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("IDutente").Value.ToString & ";" & Rec("Progressivo").Value.ToString & ";" & Rec("idPermesso").Value.ToString & ";" &
										   Rec("Descrizione").Value.ToString & ";" & Rec("NomePerCodice").Value.ToString & "§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaTuttiPermessiUtente(Squadra As String, idTipologia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""

				Try
					Sql = "SELECT A.idPermesso, B.Descrizione From [Generale].[dbo].[Permessi_Composizione] A " &
						"Left Join [Generale].[dbo].[Permessi_Lista] B " &
						"On A.idPermesso = B.idPermesso " &
						"Where Eliminato = 'N' " &
						"And A.idTipologia = " & idTipologia & " " &
						"Order By Descrizione"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							' Ritorno = StringaErrore & " Nessun permesso ritornato"
						Else
							Ritorno = ""
							Do Until Rec.Eof()
								Ritorno &= Rec("idPermesso").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

								Rec.MoveNext()
							Loop
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					'				Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaPermessiUtente(Squadra As String, IDutente As Integer, Progressivo As Integer, idTipologia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim Rec As Object
				Dim PermessoDaEliminare As String = ""

				Try
					Sql = "Select B.Descrizione From PermessiUtente A " &
						"Left Join [Generale].[dbo].[Permessi_Composizione] B On A.idPermesso = B.idPermesso " &
						"Where IDutente = " & IDutente & " And Progressivo = " & Progressivo & " And B.idTipologia = " & idTipologia
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ok = False
							Ritorno = StringaErrore & " Permesso non rilevato"
						Else
							PermessoDaEliminare = Rec(0).Value
						End If
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				If PermessoDaEliminare.Contains("/") Then
					Dim Campi() As String = PermessoDaEliminare.Split("/")
					Dim Chiave As String = Campi(0).Trim & " /"
					Dim CiSonoAltri As Boolean = False

					Try
						Sql = "Select B.Descrizione From PermessiUtente A " &
							"Left Join [Generale].[dbo].[Permessi_Composizione] B On A.idPermesso = B.idPermesso " &
							"Where IDutente = " & IDutente & " And Progressivo <> " & Progressivo & " And B.idTipologia = " & idTipologia
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ok = False
								Ritorno = StringaErrore & " Permesso non rilevato"
							Else
								Do Until Rec.Eof()
									If Rec(0).Value.ToString.Contains(Chiave) Then
										CiSonoAltri = True
										Exit Do
									End If

									Rec.MoveNext()
								Loop
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Not CiSonoAltri Then
						Dim IdPermessoPadre As Integer = -1

						Try
							Sql = "Select * From [Generale].[dbo].[Permessi_Composizione] Where Descrizione = '" & Chiave.Replace("/", "").Trim & "'"
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof() Then
									Ok = False
									Ritorno = StringaErrore & " Permesso padre non rilevato"
								Else
									IdPermessoPadre = Rec("idPermesso").Value
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try

						If IdPermessoPadre > -1 Then
							Try
								Sql = "Delete PermessiUtente Where IDutente=" & IDutente & " And idPermesso = " & IdPermessoPadre
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If
					End If
				End If

				Try
					Sql = "Delete PermessiUtente Where IDutente=" & IDutente & " AND Progressivo=" & Progressivo
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InserisciPermessiUtente(Squadra As String, IDutente As Integer, Progressivo As Integer, Permesso As Integer, idTipologia As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = New clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim Rec As Object
				Dim ProgPerm As Integer = -1

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						If TipoDB = "SQLSERVER" Then
							Sql = "SELECT IsNull(Max(Progressivo),0)+1 FROM PermessiUtente Where IDutente=" & IDutente
						Else
							Sql = "SELECT Coalesce(Max(Progressivo),0)+1 FROM PermessiUtente Where IDutente=" & IDutente
						End If
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							'If Rec(0).Value Is DBNull.Value Then
							'	ProgPerm = 1
							'Else
							ProgPerm = Rec(0).Value
							'End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					Dim DescrizionePermesso As String = ""

					Try
						Sql = "SELECT B.Descrizione FROM [Generale].[dbo].[Permessi_Composizione] A " &
							"Left Join [Generale].[dbo].[Permessi_Lista] B On A.idPermesso = B.idPermesso " &
							"Where idPermesso=" & Permesso
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ritorno = StringaErrore & " Permesso non trovato"
							Else
								DescrizionePermesso = Rec("Descrizione").Value
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Not Ritorno.Contains(StringaErrore) Then
						Try
							If Not Ritorno.Contains(StringaErrore) Then
								Sql = "Insert Into PermessiUtente Values (" &
									" " & SistemaNumero(IDutente) & "," &
									" " & SistemaNumero(ProgPerm) & "," &
									" " & SistemaNumero(Permesso) & " " &
									")"

								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If

					If DescrizionePermesso.Contains("/") Then
						Dim lPadre() As String = DescrizionePermesso.Split("/")
						Dim Padre As String = lPadre(0).Trim
						Dim idFunzionePadre As Integer = -1

						Try
							Sql = "SELECT * FROM [Generale].[dbo].[Permessi_Composizione] A " &
								"Left Join [Generale].[dbo].[Permessi_Lista] B On A.idPermesso = B.idPermesso " &
								"Where B.Descrizione = '" & Padre & "' And B.idTipologia=" & idTipologia
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Permesso padre non trovato"
								Else
									idFunzionePadre = Rec("idPermesso").Value
								End If
								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try

						Dim DaAggiungere As Boolean = False

						If Not Ritorno.Contains(StringaErrore) Then
							Try
								Sql = "SELECT * FROM PermessiUtente Where IDutente = " & IDutente & " And idPermesso = " & idFunzionePadre
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec.Eof() Then
										DaAggiungere = True
									End If
									Rec.Close()
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If

						If Not Ritorno.Contains(StringaErrore) And DaAggiungere Then
							Try
								If Not Ritorno.Contains(StringaErrore) Then
									ProgPerm += 1
									Sql = "Insert Into PermessiUtente Values (" &
										" " & SistemaNumero(IDutente) & "," &
										" " & SistemaNumero(ProgPerm) & "," &
										" " & SistemaNumero(idFunzionePadre) & " " &
										")"

									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If
					End If

					If Ritorno.Contains(StringaErrore) Then
						Dim Ritorno2 As String

						Sql = "Delete From PermessiUtente Where IDutente=" & IDutente
						Ritorno2 = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					End If
				Else
					Ok = False
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	'<WebMethod()>
	'Public Function ModificaPermessiUtente(Squadra As String, IDutente As Integer, progressivo As Integer, permesso As Integer) As String
	'	Dim Ritorno As String = ""
	'	Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

	'	If Connessione = "" Then
	'		Ritorno = ErroreConnessioneNonValida
	'	Else
	'		Dim Conn As Object = new clsGestioneDB

	'		If TypeOf (Conn) Is String Then
	'			Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
	'		Else
	'			Dim Rec as object
	'			Dim Sql As String = ""
	'			Dim Ok As Boolean = True

	'			Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
	'			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

	'			If Not Ritorno.Contains(StringaErrore) Then
	'				Sql = "Delete From PermessiUtente Where IDutente=" & IDutente & " AND progressivo=" & progressivo
	'				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'				If Ritorno.Contains(StringaErrore) Then
	'					Ok = False
	'				End If

	'				If Ok Then
	'					Sql = "Insert Into PermessiUtente Values (" &
	'							" " & SistemaNumero(IDutente) & "," &
	'							" " & SistemaNumero(progressivo) & "," &
	'							" " & SistemaNumero(permesso) & " " &
	'							")"
	'					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'					If Ritorno.Contains(StringaErrore) Then
	'						Ok = False
	'					End If
	'				End If
	'			Else
	'				Ok = False
	'			End If

	'			If Ok Then
	'				Sql = "commit"
	'				Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'			Else
	'				Sql = "rollback"
	'				Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
	'			End If

	'			Conn.Close()
	'		End If
	'	End If

	'	Return Ritorno
	'End Function


End Class