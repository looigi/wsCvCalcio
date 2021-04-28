Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://wsWidgetCVC.it/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsWidget
	Inherits System.Web.Services.WebService

	Public Class parametriConteggi
		Public Squadra As String
		Public Tutte As String
	End Class

	Public Class parametriFirme
		Public Squadra As String
		Public Tutte As String
	End Class

	Public Class parametriIscritti
		Public Squadra As String
	End Class

	Public Class parametriQuote
		Public Squadra As String
	End Class

	Public Class parametriIndicatori
		Public Squadra As String
	End Class

	<WebMethod()>
	Public Function AggiornaDati(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")
		Dim s() As String = Squadra.Split("_")
		Dim Anno As Integer = Val(s(0))
		Dim idSquadra As Integer = Val(s(1))

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")

				Sql = "Select * From AggiornamentoWidgets Where idSquadra=" & idSquadra
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						CreaConteggi(Squadra)
						CreaFirmeDaValidare(Squadra, "S")
						CreaIndicatori(Squadra)
						CreaIscritti(Squadra)
						CreaQuoteNonSaldate(Squadra)

						Sql = "Update [Generale].[dbo].[AggiornamentoWidgets] Set AggiornaWidgets='N' Where idSquadra=" & idSquadra
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno = "OK" Then
							Ritorno = "*"
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function PulisceDati(Squadra As String) As String
		'Dim thread As New Thread(AddressOf RitornaFirmeDaValidareThread)
		'Dim parameters As New parametriFirme
		'parameters.Squadra = Squadra
		'parameters.Tutte = Tutte
		'thread.Start(parameters)
		Dim s() As String = Squadra.Split("_")
		Dim Anno As Integer = Val(s(0))
		Dim idSquadra As Integer = Val(s(1))

		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")

				Sql = "Select * From AggiornamentoWidgets Where idSquadra=" & idSquadra
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Sql = "Insert Into AggiornamentoWidgets Values (" & idSquadra & ", 'S')"
					Else
						Sql = "Update AggiornamentoWidgets Set AggiornaWidgets='S' Where idSquadra=" & idSquadra
					End If
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno = "OK" Then
						Ritorno = "*"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaConteggi(Squadra As String) As String
		'Dim thread As New Thread(AddressOf RitornaConteggiThread)
		'Dim parameters As New parametriConteggi
		'parameters.Squadra = Squadra
		'thread.Start(parameters)

		Return RitornaConteggiThread(Squadra) ' "*"
	End Function

	<WebMethod()>
	Public Function RitornaConteggi(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim Trovato As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From WidgetConteggi"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Trovato = True
						Do Until Rec.Eof
							Ritorno &= Rec("idTipologia").Value & ";" & Rec("Descrizione").Value & ";" & Rec("Quanti").Value & "§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close
				End If
			End If

			If Not Trovato Then
				CreaConteggi(Squadra)
			End If
		End If

		Return Ritorno
	End Function

	Private Function RitornaConteggiThread(ByVal data As String) As String
		Dim Squadra As String = data ' .Squadra

		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)

				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Delete From WidgetConteggi"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Sql = "Insert Into WidgetConteggi Select A.idTipologia, B.Descrizione, Count(*) As Quanti From [Generale].[dbo].[Utenti] A " &
					"Left Join [Generale].[dbo].[Tipologie] B On A.idTipologia = B.idTipologia  " &
					"Where Eliminato = 'N' And B.idTipologia > 2 And idSquadra = " & codSquadra & " " &
					"Group By A.idTipologia, B.Descrizione " &
					"Order By Descrizione"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaFirmeDaValidare(Squadra As String, Tutte As String) As String
		'Dim thread As New Thread(AddressOf RitornaFirmeDaValidareThread)
		'Dim parameters As New parametriFirme
		'parameters.Squadra = Squadra
		'parameters.Tutte = Tutte
		'thread.Start(parameters)

		Return RitornaFirmeDaValidareThread(Squadra, Tutte) ' "*"
	End Function

	<WebMethod()>
	Public Function RitornaFirmeDaValidare(Squadra As String, Tutte As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim Trovato As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")

				Dim Altro As String = ""
				If Tutte = "" Or Tutte = "N" Or Tutte = "NO" Then
					Altro = "Top 3"
				End If

				Dim Sql As String = "Select " & Altro & " * From WidgetFirme"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Trovato = True
						Do Until Rec.Eof
							Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("idGenitore").Value.ToString & ";" &
									Rec("Datella").Value.ToString.Trim & ";" &
									Rec("DataFirma").Value.ToString.Trim & ";" &
									Rec("Giocatore").Value.ToString.Trim & ";" &
									Rec("Genitore").Value.ToString.Trim & ";" &
									"§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close
				End If
			End If

			If Not Trovato Then
				CreaFirmeDaValidare(Squadra, Tutte)
			End If
		End If

		Return Ritorno
	End Function

	Private Function RitornaFirmeDaValidareThread(ByVal Data As String, Data2 As String) As String
		Dim Squadra As String = Data ' .Squadra
		Dim Tutte As String = Data2 ' .Tutte

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
				Dim Altro As String = ""

				If Tutte = "" Or Tutte = "N" Or Tutte = "NO" Then
					Altro = "Top 3"
				End If

				Dim Sql As String = "Delete From WidgetFirme"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Sql = "Insert Into WidgetFirme Select " & Altro & " A.*, B.Cognome + ' ' + B.Nome As Giocatore, " &
					"CASE A.idGenitore " &
					"     WHEN 1 THEN C.Genitore1 " &
					"     WHEN 2 THEN C.Genitore2 " &
					"     WHEN 3 THEN B.Cognome + ' ' + B.Nome " &
					"END As Genitore " &
					"From GiocatoriFirme A " &
					"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
					"Left Join GiocatoriDettaglio C On A.idGiocatore = C.idGiocatore " &
					"Where (DataFirma Is Not Null And DataFirma <> '') And (Validazione Is Null Or Validazione = '') And idGenitore < 100"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaIscritti(Squadra As String) As String
		'Dim thread As New Thread(AddressOf RitornaIscrittiThread)
		'Dim parameters As New parametriIscritti
		'parameters.Squadra = Squadra
		'thread.Start(parameters)


		Return RitornaIscrittiThread(Squadra) ' "*"
	End Function

	<WebMethod()>
	Public Function RitornaIscritti(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim Trovato1 As Boolean = False
			Dim Trovato2 As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				'Dim Tutti As Integer = 0
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				'Sql = "Select * From WidgetIscritti1"
				'Rec = LeggeQuery(Conn, Sql, Connessione)
				'If TypeOf (Rec) Is String Then
				'	Ritorno = Rec
				'Else
				'	If Not Rec.Eof Then
				'		Trovato1 = True
				'		If Rec(0).Value Is DBNull.Value Then
				'			Tutti = 0
				'		Else
				'			Tutti = Rec(0).Value
				'		End If
				'		Rec.Close
				'	End If
				'End If

				Sql = "Select * From WidgetIscritti2 Order By Anno"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Trovato2 = True
						Do Until Rec.Eof
							Ritorno &= Rec("Anno").Value & ";" & Rec("Anno").Value & ";" & Rec("Quanti").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()

						'Ritorno &= "-1;Tutti;" & Tutti & "§"
					End If
				End If
			End If

			If Not Trovato1 Or Not Trovato2 Then
				CreaIscritti(Squadra)
			End If
		End If

		Return Ritorno
	End Function

	Private Function RitornaIscrittiThread(ByVal Data As String) As String
		Dim Squadra As String = Data '.Squadra

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
				Dim listaCategorie As New List(Of String)
				Dim idCategorie As New List(Of String)
				Dim Ok As Boolean = True

				Dim Sql As String = "Delete From WidgetIscritti1"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Dim Tutti As Integer = 0
				Sql = "Insert Into WidgetIscritti1 Select Count(*) From Giocatori Where Eliminato='N'" '  And RapportoCompleto='S'"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Ok Then
					Sql = "Delete From WidgetIscritti2"
					Ritorno = EsegueSql(Conn, Sql, Connessione)

					Sql = "Insert Into WidgetIscritti2 Select YEAR(CONVERT(Date, DataDiNascita)) As Anno, Count(*) As Quanti From Giocatori " &
						"Where Eliminato = 'N' " & ' And RapportoCompleto='S' " &
						"Group By YEAR(CONVERT(date, DataDiNascita)) " &
						"Order By 1"
					Ritorno = EsegueSql(Conn, Sql, Connessione)
				End If
			End If

			Conn.Close()
			'End If
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaQuoteNonSaldate(Squadra As String) As String
		'Dim thread As New Thread(AddressOf RitornaQuoteNonSaldateThread)
		'Dim parameters As New parametriQuote
		'parameters.Squadra = Squadra
		'thread.Start(parameters)

		Return RitornaQuoteNonSaldateThread(Squadra) ' "*"
	End Function

	<WebMethod()>
	Public Function RitornaQuoteNonSaldate(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim Trovato As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From WidgetQuoteNonSaldate Order By Anno1, Anno2"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Trovato = True
						Ritorno = ""
						Do Until Rec.Eof
							Ritorno &= Rec("Anno1").Value & ";" & Rec("Anno2").Value & ";" & Rec("Differenza").Value & "§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close

					Dim Totalone As String = "0"

					Sql = "Select * From WidgetTotaleQuote"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof Then
							Trovato = True
							Totalone = Rec("Totale").Value

							Rec.close
						End If
					End If

					Ritorno = Totalone & "|" & Ritorno
				End If
			End If

			If Not Trovato Then
				CreaQuoteNonSaldate(Squadra)
			End If
		End If

		Return Ritorno
	End Function

	Private Function RitornaQuoteNonSaldateThread(ByVal Data As String) As String
		Dim Squadra As String = Data ' .Squadra

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
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Rec3 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String
				'Dim listaCategorie As New List(Of String)
				'Dim idCategorie As New List(Of String)
				Dim Ok As Boolean = True

				'Sql = "Select * From Categorie Where Eliminato='N' "
				'Rec = LeggeQuery(Conn, Sql, Connessione)
				'If TypeOf (Rec) Is String Then
				'	Ritorno = Rec
				'	Ok = False
				'Else
				'	Do Until Rec.Eof
				'		idCategorie.Add(Rec("idCategoria").Value)
				'		listaCategorie.Add(Rec("Descrizione").Value)

				'		Rec.MoveNext
				'	Loop
				'	Rec.Close
				'End If

				'If Ok Then
				'Dim Differenza(idCategorie.Count) As Single
				Sql = "Delete From WidgetTotaleQuote"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Dim Totalone As String = "0"

				Sql = "Select Sum(DaPagare) - Sum(Sconto) From ( " &
					"Select Cognome, Nome, DataDiNascita, Descrizione, DaPagare, Sconto, TotalePagato, (DaPagare - Sconto) - TotalePagato As Differenza From (  " &
					"Select Cognome, Nome, DataDiNascita, C.Descrizione, IsNull(C.Importo, 0) As DaPagare, IsNull(B.Sconto, 0) As Sconto, " &
					"(Select IsNull(Sum(Pagamento),0) From GiocatoriPagamenti Where idGiocatore=A.idGiocatore And Eliminato='N' And idTipoPagamento = 1 And Validato='S') As TotalePagato " &
					"From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Left Join Quote C On B.idQuota = C.idQuota " &
					"Where A.Eliminato='N' And C.Eliminato='N' " &
					") As A ) as B"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec(0).Value Is DBNull.Value Then
						Totalone = 0
					Else
						Totalone = Rec(0).Value
					End If
				End If

				Sql = "Insert Into WidgetTotaleQuote Values (" & Totalone.Replace(",", ".") & ")"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Sql = "Delete From WidgetQuoteNonSaldate"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				'For Each id As String In idCategorie
				Sql = "Select YEAR(CONVERT(date, DataDiNascita)) As Anno From Giocatori " &
					"Where Eliminato = 'N' " &
					"Group By YEAR(CONVERT(date, DataDiNascita)) " &
					"Order By 1"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					Do Until Rec.Eof
						Sql = "Select A.idGiocatore, TotalePagamento, B.Sconto From Giocatori A " &
							"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
							"Where A.Eliminato = 'N' And YEAR(CONVERT(date, DataDiNascita)) = " & Rec("Anno").Value
						Rec2 = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec2) Is String Then
							Ritorno = Rec2
							Ok = False
						Else
							Dim Differenza As Single = 0

							Do Until Rec2.Eof
								Dim s As String = ("" & Rec2("TotalePagamento").value).replace(",", ".")
								Dim s2 As String = ("" & Rec2("Sconto").value).replace(",", ".")
								Dim TotalePagamento As Single = Val(s) - Val(s2)
								Dim Pagato As Single = 0

								Sql = "Select Sum(Pagamento) From GiocatoriPagamenti Where idGiocatore=" & Rec2("idGiocatore").Value & " And Eliminato='N' And idTipoPagamento=1 And Validato='S'"
								Rec3 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec3) Is String Then
									Ritorno = Rec3
									Ok = False
								Else
									If Rec3(0).Value Is DBNull.Value Then
										Pagato = 0
									Else
										Dim p As String = ("" & Rec3(0).Value).replace(",", ".")
										Pagato = Val(p)
									End If
									Rec3.Close
								End If
								Differenza += (TotalePagamento - Pagato)

								Rec2.MoveNext
							Loop
							Rec2.Close

							'Ritorno &= Rec("Anno").Value & ";" & Rec("Anno").Value & ";" & Differenza & "§"

							Sql = "Insert Into WidgetQuoteNonSaldate Values (" & Rec("Anno").Value & ", " & Rec("Anno").Value & ", " & Differenza & ")"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
						End If

						'Dim qualeCat As Integer = 0
						'For Each idcat As String In idCategorie
						'	If Val(idcat) = Val(id) Then
						'		Differenza(qualeCat) += (TotalePagamento - Pagato)
						'		Exit For
						'	End If
						'	qualeCat += 1
						'Next

						Rec.MoveNext
					Loop
					Rec.Close
				End If
				'Next

				'Sql = "Select Sum(C.TotalePagamento) - Sum(Pagamento) From GiocatoriPagamenti A " &
				'	"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
				'	"Left Join GiocatoriDettaglio C On A.idGiocatore = C.idGiocatore " &
				'	"Where B.Categorie = '' And A.Eliminato='N'"
				'Rec = LeggeQuery(Conn, Sql, Connessione)
				'If TypeOf (Rec) Is String Then
				'	Ritorno = Rec
				'	Ok = False
				'Else
				'	If Rec(0).Value Is DBNull.Value Then
				'		Ritorno &= "-1;Nessuna Categoria;0§"
				'	Else
				'		Ritorno &= "-1;Nessuna Categoria;" & Rec(0).Value & "§"
				'	End If
				'	Rec.Close
				'End If

				'Dim quale As Integer = 0

				'For Each categ As String In listaCategorie
				'	Dim d As String = (Int(Differenza(quale) * 100) / 100).ToString
				'	Ritorno &= idCategorie.Item(quale) & ";" & categ & ";" & d & "§"

				'	quale += 1
				'Next
			End If
			'End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaProssimiEventi(Squadra As String, Limite As String) As String
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
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String
				Dim Altro As String = ""

				If Limite <> "" Then
					Altro = "Top " & Limite
				End If

				Sql = "Select " & Altro & " * From ( " &
					"Select 'Cert. Scad.' As Cosa, A.idGiocatore As Id, 'Certificato medico scaduto' As PrimoCampo, A.Cognome + ' ' + A.Nome As SecondoCampo, B.ScadenzaCertificatoMedico As Data From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Where A.Eliminato='N' And B.ScadenzaCertificatoMedico Is Not Null And B.ScadenzaCertificatoMedico <> '' " &
					"And Convert(DateTime, B.ScadenzaCertificatoMedico, 121) < CURRENT_TIMESTAMP " &
					"Union All " &
					"Select 'Cert. Med.' As Cosa, A.idGiocatore As Id, A.Cognome As PrimoCampo, A.Nome As SecondoCampo, CONVERT(date, B.ScadenzaCertificatoMedico) As Data From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Where A.Eliminato='N' And CertificatoMedico = 'S' And A.Eliminato = 'N' And Convert(DateTime, B.ScadenzaCertificatoMedico ,121) <= DateAdd(Day, 30, CURRENT_TIMESTAMP) " &
					"Union All " &
					"Select 'Partita' As Cosa, idPartita As Id, B.Descrizione As PrimoCampo, C.Descrizione As SecondoCampo, CONVERT(date, DataOra) As Data From Partite A " &
					"Left Join Categorie B On A.idCategoria = B.idCategoria " &
					"Left Join SquadreAvversarie C On A.idAvversario = C.idAvversario " &
					"Union All " &
					"Select 'Evento' As Cosa, idEvento As Id, Titolo As PrimoCampo, '' As SecondoCampo, CONVERT(date, Inizio) As Data From EventiConvocazioni " &
					"Where idTipologia = 2) A  " &
					"Where Data > GETDATE() Or Cosa = 'Cert. Scad.' " &
					"Order By Data"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Ritorno = ""
					Do Until Rec.Eof
						Ritorno &= Rec("Cosa").Value & ";" & Rec("Id").Value & ";" & Rec("PrimoCampo").Value & ";" & Rec("SecondoCampo").Value & ";" & Rec("Data").Value & "§"

						Rec.MoveNext
					Loop
					Rec.Close
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaIndicatori(Squadra As String) As String
		'Dim thread As New Thread(AddressOf RitornaIndicatoriThread)
		'Dim parameters As New parametriIndicatori
		'parameters.Squadra = Squadra
		'thread.Start(parameters)

		Return RitornaIndicatoriThread(Squadra)
	End Function

	Private Function RitornaIndicatoriThread(ByVal Data As String) As String
		Dim Squadra As String = Data ' .Squadra

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
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")

				Sql = "Delete From WidgetIndicatori"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Dim SenzaQuota As Integer = 0
				Dim CertificatoScadutoAssente As Integer = 0
				Dim SenzaFirma As Integer = 0
				Dim KitNonConsegnato As Integer = 0

				' Giocatori senza quota
				Sql = "Select Count(*) " &
					"From " &
					"Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Left Join Quote C On B.idQuota = C.idQuota And C.Eliminato = 'N' " &
					"Where A.Eliminato = 'N' And C.Descrizione Is Null"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Return Ritorno
				Else
					If Rec(0).Value Is DBNull.Value Then
						SenzaQuota = 0
					Else
						SenzaQuota = Rec(0).Value
					End If
					Rec.Close
				End If

				' Certificato scaduto / assente
				Sql = "Select Count(*) From Giocatori A " &
						"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						"Where A.Eliminato = 'N' And "
				Sql &= " ((B.ScadenzaCertificatoMedico Is Not Null And B.ScadenzaCertificatoMedico <> '' And Convert(DateTime, B.ScadenzaCertificatoMedico ,121) <= CURRENT_TIMESTAMP And B.CertificatoMedico = 'S') Or "
				Sql &= " (B.CertificatoMedico Is Null Or B.CertificatoMedico = '' Or B.CertificatoMedico = 'N' Or (B.CertificatoMedico = 'S' And (B.ScadenzaCertificatoMedico Is Null Or B.ScadenzaCertificatoMedico = ''))))"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Return Ritorno
				Else
					If Rec(0).Value Is DBNull.Value Then
						CertificatoScadutoAssente = 0
					Else
						CertificatoScadutoAssente = Rec(0).Value
					End If
					Rec.Close
				End If

				' Giocatori senza firma
				Dim gf As New GestioneFilesDirectory
				Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
				pp = pp.Replace(vbCrLf, "")
				pp = pp.Trim()
				If Strings.Right(pp, 1) = "\" Then
					pp = Mid(pp, 1, pp.Length - 1)
				End If
				Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim P() As String = PathAllegati.Split(";")
				P(2) = P(2).Replace(vbCrLf, "").Trim
				If Strings.Right(P(2), 1) = "\" Then
					P(2) = Mid(P(2), 1, P(2).Length - 1)
				End If
				Dim CodSquadra() As String = Squadra.Split("_")
				Dim idSquadra As Integer = Val(CodSquadra(1))
				Dim idAnno As String = Val(CodSquadra(0)).ToString.Trim

				Dim NomeSquadra As String = ""
				Dim IscrFirmaEntrambi As String = ""

				Sql = "Select * From Anni"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Return Ritorno
				Else
					If Rec.Eof = False Then
						IscrFirmaEntrambi = Rec("iscrFirmaEntrambi").Value
					End If
				End If

				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & idSquadra
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Return Ritorno
				Else
					If Rec.Eof = False Then
						NomeSquadra = Rec("Descrizione").Value
					End If
				End If

				Sql = "Select * From Giocatori A " &
						"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						"Where A.Eliminato = 'N' "
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Return Ritorno
				Else
					Do Until Rec.Eof
						Dim urlFirma As String = ""
						Dim CiSonoFirme As Boolean = True

						If "" & Rec("Maggiorenne").Value = "S" Then
							If "" & Rec("AbilitaFirmaGenitore3").Value = "S" Then
								urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_3.kgb"
								If Not File.Exists(urlFirma) Then
									CiSonoFirme = False
								End If
							Else
								If "" & Rec("FirmaAnalogicaGenitore3").Value = "N" Then
									CiSonoFirme = False
								End If
							End If
						Else
							If "" & Rec("GenitoriSeparati").Value = "S" Then
								If "" & Rec("AffidamentoCongiunto").Value = "S" Then
									If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
										If Not File.Exists(urlFirma) Then
											CiSonoFirme = False
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
											CiSonoFirme = False
										End If
									End If

									If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
										If Not File.Exists(urlFirma) Then
											CiSonoFirme = False
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
											CiSonoFirme = False
										End If
									End If
								Else
									If "" & Rec("idTutore").Value = "1" Then
										If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
											urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
											If Not File.Exists(urlFirma) Then
												CiSonoFirme = False
											End If
										Else
											If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
												CiSonoFirme = False
											End If
										End If
									Else
										If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
											urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
											If Not File.Exists(urlFirma) Then
												CiSonoFirme = False
											End If
										Else
											If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
												CiSonoFirme = False
											End If
										End If
									End If
								End If
							Else
								If IscrFirmaEntrambi = "S" Then
									If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
										If Not File.Exists(urlFirma) Then
											CiSonoFirme = False
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
											CiSonoFirme = False
										End If
									End If

									If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
										If Not File.Exists(urlFirma) Then
											CiSonoFirme = False
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
											CiSonoFirme = False
										End If
									End If
								Else
									If "" & Rec("Genitore1").Value <> "" Then
										If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
											urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
											If Not File.Exists(urlFirma) Then
												CiSonoFirme = False
											End If
										Else
											If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
												CiSonoFirme = False
											End If
										End If
									Else
										If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
											urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
											If Not File.Exists(urlFirma) Then
												CiSonoFirme = False
											End If
										Else
											If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
												CiSonoFirme = False
											End If
										End If
									End If
								End If
							End If
						End If

						If Not CiSonoFirme Then
							SenzaFirma += 1
						End If

						Dim NomeKit As String = ""
						Dim TagliaKit As String = ""

						Sql = "Select C.Quantita, QuantitaConsegnata, D.Descrizione As NomeKit, F.Descrizione As Taglia, G.Descrizione As Elemento From KitComposizione A " &
							"Left Join KitGiocatori B On A.idTipoKit = B.idTipokit And A.idElemento = B.idElemento " &
							"Left Join KitComposizione C On B.idTipoKit = C.idTipoKit And B.idElemento = C.idElemento " &
							"Left Join KitTipologie D On D.idTipoKit = C.idTipoKit " &
							"Left Join Giocatori E On B.idGiocatore = E.idGiocatore " &
							"Left Join Taglie F On E.idTaglia = F.idTaglia " &
							"Left Join KitElementi G On G.idElemento = C.idElemento " &
							"Where B.idGiocatore = " & Rec("idGiocatore").Value & " And C.Eliminato='N' And A.Eliminato='N' And D.Eliminato='N' And E.Eliminato='N' And G.Eliminato='N'"
						Rec2 = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec2) Is String Then
							Ritorno = Rec2
							Return Ritorno
						Else
							Dim Tutto As Boolean = True
							Dim Qualcosa As Boolean = False

							If Rec2.eof Then
								Tutto = False
							Else
								Do Until Rec2.Eof()
									If NomeKit = "" Then
										NomeKit = "" & Rec2("NomeKit").Value
										TagliaKit = "" & Rec2("Taglia").Value
									End If

									If Val(Rec2("QuantitaConsegnata").Value) > 0 Then
										Qualcosa = True
										If Val(Rec2("QuantitaConsegnata").Value) < Val(Rec2("Quantita").Value) Then
											Tutto = False
										End If
									Else
										Tutto = False
									End If

									Rec2.MoveNext()
								Loop
								Rec2.Close()
							End If

							' Kit consegnato No
							Dim Preso As Boolean = False

							If Tutto = False Then
								KitNonConsegnato += 1
							End If
						End If

						Rec.MoveNext()
					Loop
				End If
				Rec.Close()

				Sql = "Insert Into WidgetIndicatori Values (" & SenzaQuota & ", " & CertificatoScadutoAssente & ", " & SenzaFirma & ", " & KitNonConsegnato & ")"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaIndicatori(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)
			Dim Trovato As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From WidgetIndicatori"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Trovato = True
						Ritorno &= Rec("SenzaQuota").Value & ";" & Rec("CertificatoScadutoAssente").Value & ";" & Rec("SenzaFirma").Value & ";" & Rec("KitNonConsegnato").Value
						Rec.Close
					End If
				End If
			End If

			If Not Trovato Then
				CreaIndicatori(Squadra)
			End If
		End If

		Return Ritorno
	End Function
End Class