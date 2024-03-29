﻿Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports ADODB

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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object

				Sql = "Select * From AggiornamentoWidgets Where idSquadra=" & idSquadra
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						CreaConteggi(Connessione, Conn, Squadra)
						CreaFirmeDaValidare(Squadra, "S")
						CreaIndicatori(Squadra)
						CreaIscritti(Squadra)
						CreaQuoteNonSaldate(Squadra)

						Sql = "Update [Generale].[dbo].[AggiornamentoWidgets] Set AggiornaWidgets='N' Where idSquadra=" & idSquadra
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object

				Sql = "Select * From AggiornamentoWidgets Where idSquadra=" & idSquadra
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Sql = "Insert Into AggiornamentoWidgets Values (" & idSquadra & ", 'S')"
					Else
						Sql = "Update AggiornamentoWidgets Set AggiornaWidgets='S' Where idSquadra=" & idSquadra
					End If
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno = "OK" Then
						Ritorno = "*"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaConteggi(Connessione As String, conn As Object, Squadra As String) As String
		'Dim thread As New Thread(AddressOf RitornaConteggiThread)
		'Dim parameters As New parametriConteggi
		'parameters.Squadra = Squadra
		'thread.Start(parameters)

		Return RitornaConteggiThread(Connessione, conn, Squadra) ' "*"
	End Function

	<WebMethod()>
	Public Function RitornaConteggi(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim Trovato As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From WidgetConteggi"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						Trovato = True
						Do Until Rec.Eof()
							Ritorno &= Rec("idTipologia").Value & ";" & Rec("Descrizione").Value & ";" & Rec("Quanti").Value & "§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close()
				End If
			End If

			If Not Trovato Then
				CreaConteggi(Connessione, Conn, Squadra)
			End If
		End If

		Return Ritorno
	End Function

	Private Function RitornaConteggiThread(Connessione As String, Conn As Object, ByVal Squadra As String) As String
		Dim Ritorno As String = ""
		Dim c() As String = Squadra.Split("_")
		Dim Anno As String = Str(Val(c(0))).Trim
		Dim codSquadra As String = c(1)

		Dim Sql As String = "Delete From WidgetConteggi"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

		Sql = "Insert Into WidgetConteggi Select A.idTipologia, B.Descrizione, " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " As Quanti From [Generale].[dbo].[Utenti] A " &
			"Left Join [Generale].[dbo].[Tipologie] B On A.idTipologia = B.idTipologia  " &
			"Where Eliminato = 'N' And B.idTipologia > 2 And idSquadra = " & codSquadra & " " &
			"Group By A.idTipologia, B.Descrizione " &
			"Order By Descrizione"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

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

		CreaFirmeDaValidare(Squadra, Tutte)

		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim Trovato As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object

				Dim Altro As String = ""
				Dim Altro2 As String = ""
				If Tutte = "" Or Tutte = "N" Or Tutte = "NO" Then
					Altro = IIf(TipoDB = "SQLSERVER", "Top 5", "")
					Altro2 = IIf(TipoDB = "SQLSERVER", "", "Limit 5")
				End If

				Dim Sql As String = "Select " & Altro & " * From WidgetFirme " & Altro2
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						Trovato = True
						Do Until Rec.Eof()
							Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("idGenitore").Value.ToString & ";" &
									Rec("Datella").Value.ToString.Trim & ";" &
									Rec("DataFirma").Value.ToString.Trim & ";" &
									Rec("Giocatore").Value.ToString.Trim & ";" &
									Rec("Genitore").Value.ToString.Trim & ";" &
									Rec("QualeFirma").Value.ToString.Trim & ";" &
									"§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close()
				End If
			End If

			'If Not Trovato Then
			'	CreaFirmeDaValidare(Squadra, Tutte)
			'End If
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Altro As String = ""
				Dim Altro2 As String = ""

				If Tutte = "" Or Tutte = "N" Or Tutte = "NO" Then
					If TipoDB = "SQLSERVER" Then
						Altro = "Top 3"
						Altro2 = ""
					Else
						Altro = ""
						Altro2 = "Limit 3"
					End If
				End If

				Dim Sql As String = "Delete From widgetfirme"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Sql = "Insert Into widgetfirme (idGiocatore, idGenitore, Datella, DataFirma, Validazione, QualeFirma, Giocatore, Genitore) " &
					"Select " & Altro & " A.*, " & IIf(TipoDB = "SQLSERVER", "B.Cognome + ' ' + B.Nome", "CONCAT(B.Cognome, ' ', B.Nome)") & " As Giocatore, " &
					"CASE A.idGenitore " &
					"     WHEN 1 THEN C.Genitore1 " &
					"     WHEN 2 THEN C.Genitore2 " &
					"     WHEN 3 THEN B.Cognome + ' ' + B.Nome " &
					"END As Genitore " &
					"From giocatorifirme A " &
					"Left Join giocatori B On A.idGiocatore = B.idGiocatore " &
					"Left Join giocatoridettaglio C On A.idGiocatore = C.idGiocatore " &
					"Where (DataFirma Is Not Null And DataFirma <> '') And (Validazione Is Null Or Validazione = '') And idGenitore < 100 " & Altro2
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim Trovato1 As Boolean = False
			Dim Trovato2 As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				'Dim Tutti As Integer = 0
				Dim Rec As Object
				Dim Sql As String = ""
				'Sql = "Select * From WidgetIscritti1"
				'Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				'If TypeOf (Rec) Is String Then
				'	Ritorno = Rec
				'Else
				'	If Not Rec.Eof() Then
				'		Trovato1 = True
				'		If Rec(0).Value Is DBNull.Value Then
				'			Tutti = 0
				'		Else
				'			Tutti = Rec(0).Value
				'		End If
				'		Rec.Close()
				'	End If
				'End If

				Sql = "Select * From WidgetIscritti2 Order By Anno"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						Trovato2 = True
						Do Until Rec.Eof()
							Ritorno &= Rec("Anno").Value & ";" & Rec("Anno").Value & ";" & Rec("Quanti").Value & "§"

							Rec.MoveNext()
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim listaCategorie As New List(Of String)
				Dim idCategorie As New List(Of String)
				Dim Ok As Boolean = True

				Dim Sql As String = "Delete From widgetiscritti1"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Dim Tutti As Integer = 0
				Sql = "Insert Into widgetiscritti1 Select " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " From Giocatori Where Eliminato='N'" '  And RapportoCompleto='S'"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)

				If Ok Then
					Sql = "Delete From widgetiscritti2"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

					Sql = "Insert Into widgetiscritti2 Select YEAR(" & IIf(TipoDB = "SQLSERVER", "CONVERT(Date, DataDiNascita)", "CONVERT(DataDiNascita, Date)") & ") As Anno, " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " As Quanti From giocatori " &
						"Where Eliminato = 'N' " & ' And RapportoCompleto='S' " &
						"Group By YEAR(" & IIf(TipoDB = "SQLSERVER", "CONVERT(date, DataDiNascita)", "CONVERT(DataDiNascita, date)") & ") " &
						"Order By 1"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)
			Dim Trovato As Boolean = False

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From WidgetQuoteNonSaldate Order By Anno1, Anno2"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						Trovato = True
						Ritorno = ""
						Do Until Rec.Eof()
							Ritorno &= Rec("Anno1").Value & ";" & Rec("Anno2").Value & ";" & Rec("Differenza").Value & "§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close()

					Dim Totalone As String = "0"

					Sql = "Select * From WidgetTotaleQuote"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof() Then
							Trovato = True
							Totalone = Rec("Totale").Value

							Rec.Close()
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Rec2 As Object
				Dim Rec3 As Object
				Dim Sql As String
				'Dim listaCategorie As New List(Of String)
				'Dim idCategorie As New List(Of String)
				Dim Ok As Boolean = True

				'Sql = "Select * From Categorie Where Eliminato='N' "
				'Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				'If TypeOf (Rec) Is String Then
				'	Ritorno = Rec
				'	Ok = False
				'Else
				'	Do Until Rec.Eof()
				'		idCategorie.Add(Rec("idCategoria").Value)
				'		listaCategorie.Add(Rec("Descrizione").Value)

				'		Rec.MoveNext
				'	Loop
				'	Rec.Close()
				'End If

				'If Ok Then
				'Dim Differenza(idCategorie.Count) As Single
				Sql = "Delete From WidgetTotaleQuote"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Dim Totalone As String = "0"

				Sql = "Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Sum(DaPagare),0) - IsNull(Sum(Sconto),0)", "Coalesce(Sum(DaPagare),0) - Coalesce(Sum(Sconto),0)") & " From ( " &
					"Select Cognome, Nome, DataDiNascita, Descrizione, DaPagare, Sconto, TotalePagato, (DaPagare - Sconto) - TotalePagato As Differenza From (  " &
					"Select Cognome, Nome, DataDiNascita, C.Descrizione, " & IIf(TipoDB = "SQLSERVER", "IsNull(C.Importo, 0)", "COALESCE(C.Importo, 0)") & " As DaPagare, " & IIf(TipoDB = "SQLSERVER", "IsNull(B.Sconto, 0)", "COALESCE(B.Sconto, 0)") & " As Sconto, " &
					"(Select " & IIf(TipoDB = "SQLSERVER", "IsNull(Sum(Pagamento),0)", "COALESCE(Sum(Pagamento),0)") & " From GiocatoriPagamenti Where idGiocatore=A.idGiocatore And Eliminato='N' And idTipoPagamento = 1 And Validato='S') As TotalePagato " &
					"From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Left Join Quote C On B.idQuota = C.idQuota " &
					"Where A.Eliminato='N' And C.Eliminato='N' " &
					") As A ) as B"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	Totalone = 0
					'Else
					Totalone = Rec(0).Value
					'End If
				End If

				Sql = "Insert Into WidgetTotaleQuote Values (" & Totalone.Replace(",", ".") & ")"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Sql = "Delete From WidgetQuoteNonSaldate"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				'For Each id As String In idCategorie
				Sql = "Select YEAR(" & IIf(TipoDB = "SQLSERVER", "CONVERT(date, DataDiNascita)", "CONVERT(DataDiNascita, date)") & ") As Anno From Giocatori " &
					"Where Eliminato = 'N' " &
					"Group By YEAR(" & IIf(TipoDB = "SQLSERVER", "CONVERT(date, DataDiNascita)", "CONVERT(DataDiNascita, date)") & ") " &
					"Order By 1"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					Do Until Rec.Eof()
						Sql = "Select A.idGiocatore, TotalePagamento, B.Sconto From Giocatori A " &
							"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
							"Where A.Eliminato = 'N' And YEAR(" & IIf(TipoDB = "SQLSERVER", "CONVERT(date, DataDiNascita)", "CONVERT(DataDiNascita, date)") & ") = " & Rec("Anno").Value
						Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
						If TypeOf (Rec2) Is String Then
							Ritorno = Rec2
							Ok = False
						Else
							Dim Differenza As Single = 0

							Do Until Rec2.Eof()
								Dim s As String = ("" & Rec2("TotalePagamento").Value).replace(",", ".")
								Dim s2 As String = ("" & Rec2("Sconto").Value).replace(",", ".")
								Dim TotalePagamento As Single = Val(s) - Val(s2)
								Dim Pagato As Single = 0

								If TipoDB = "SQLSERVER" Then
									Sql = "Select IsNull(Sum(Pagamento),0) From GiocatoriPagamenti Where idGiocatore=" & Rec2("idGiocatore").Value & " And Eliminato='N' And idTipoPagamento=1 And Validato='S'"
								Else
									Sql = "Select Coalesce(Sum(Pagamento),0) From GiocatoriPagamenti Where idGiocatore=" & Rec2("idGiocatore").Value & " And Eliminato='N' And idTipoPagamento=1 And Validato='S'"
								End If
								Rec3 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec3) Is String Then
									Ritorno = Rec3
									Ok = False
								Else
									'If Rec3(0).Value Is DBNull.Value Then
									'	Pagato = 0
									'Else
									Dim p As String = ("" & Rec3(0).Value).replace(",", ".")
									Pagato = Val(p)
									'End If
									Rec3.Close()
								End If
								Differenza += (TotalePagamento - Pagato)

								Rec2.MoveNext()
							Loop
							Rec2.Close()

							'Ritorno &= Rec("Anno").Value & ";" & Rec("Anno").Value & ";" & Differenza & "§"

							Sql = "Insert Into WidgetQuoteNonSaldate Values (" & Rec("Anno").Value & ", " & Rec("Anno").Value & ", " & Differenza & ")"
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						End If

						'Dim qualeCat As Integer = 0
						'For Each idcat As String In idCategorie
						'	If Val(idcat) = Val(id) Then
						'		Differenza(qualeCat) += (TotalePagamento - Pagato)
						'		Exit For
						'	End If
						'	qualeCat += 1
						'Next

						Rec.MoveNext()
					Loop
					Rec.Close()
				End If
				'Next

				'Sql = "Select Sum(C.TotalePagamento) - Sum(Pagamento) From GiocatoriPagamenti A " &
				'	"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
				'	"Left Join GiocatoriDettaglio C On A.idGiocatore = C.idGiocatore " &
				'	"Where B.Categorie = '' And A.Eliminato='N'"
				'Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				'If TypeOf (Rec) Is String Then
				'	Ritorno = Rec
				'	Ok = False
				'Else
				'	If Rec(0).Value Is DBNull.Value Then
				'		Ritorno &= "-1;Nessuna Categoria;0§"
				'	Else
				'		Ritorno &= "-1;Nessuna Categoria;" & Rec(0).Value & "§"
				'	End If
				'	Rec.Close()
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Rec2 As Object
				Dim Sql As String
				Dim Altro As String = ""
				Dim Altro2 As String = ""

				If Limite <> "" Then
					If TipoDB = "SQLSERVER" Then
						Altro = "Top " & Limite
						Altro2 = ""
					Else
						Altro = ""
						Altro2 = "Limit " & Limite
					End If
				End If

				Sql = "Select * From (Select " & Altro & " * From ( " &
					"Select 'Cert. Scad.' As Cosa, A.idGiocatore As Id, 'Certificato medico scaduto' As PrimoCampo, " & IIf(TipoDB = "SQLSERVER", "A.Cognome + ' ' + A.Nome", "Concat(A.Cognome, ' ', A.Nome)") & " As SecondoCampo, B.ScadenzaCertificatoMedico As Data From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Where A.Eliminato='N' And B.ScadenzaCertificatoMedico Is Not Null And B.ScadenzaCertificatoMedico <> '' " &
					"And " & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico, 121)", "Convert(B.ScadenzaCertificatoMedico, DateTime)") & " < " & IIf(TipoDB = "SQLSERVER", "CURRENT_TIMESTAMP", "SUBDATE(NOW(), INTERVAL 1 DAY)") & " " &
					"Union All " &
					"Select 'Cert. Med.' As Cosa, A.idGiocatore As Id, A.Cognome As PrimoCampo, A.Nome As SecondoCampo, " & IIf(TipoDB = "SQLSERVER", "CONVERT(date, B.ScadenzaCertificatoMedico)", "CONVERT(B.ScadenzaCertificatoMedico, date)") & "As Data From Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Where A.Eliminato='N' And CertificatoMedico = 'S' And A.Eliminato = 'N' And " & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico ,121)", "Convert(B.ScadenzaCertificatoMedico, DateTime)") & " <= " & IIf(TipoDB = "SQLSERVER", "DateAdd(Day, 30, CURRENT_TIMESTAMP)", "ADDDATE(CURRENT_TIMESTAMP, 30)") & " " &
					"Union All " &
					"Select 'Partita' As Cosa, idPartita As Id, B.Descrizione As PrimoCampo, C.Descrizione As SecondoCampo, " & IIf(TipoDB = "SQLSERVER", "CONVERT(date, DataOra)", "CONVERT(DataOra, date)") & " As Data From Partite A " &
					"Left Join Categorie B On A.idCategoria = B.idCategoria " &
					"Left Join SquadreAvversarie C On A.idAvversario = C.idAvversario " &
					"Union All " &
					"Select 'Evento' As Cosa, idEvento As Id, Titolo As PrimoCampo, '' As SecondoCampo, " & IIf(TipoDB = "SQLSERVER", "CONVERT(date, Inizio)", "CONVERT(Inizio, date)") & " As Data From EventiConvocazioni " &
					"Where idTipologia = 2 " &
					") A  " &
					"Where Data > " & IIf(TipoDB = "SQLSERVER", "GETDATE()", "CURRENT_DATE()") & " Or Cosa = 'Cert. Scad.' " &
					"Order By Data " & Altro2 & " " &
					") As B " &
					"Union All " &
					"Select 'Compleanno' As Cosa, '' As Id, " & IIf(TipoDB = "SQLSERVER", "Cognome + ' ' + Nome", "concat(cognome, ' ',  nome)") & " As PrimoCampo, Categorie As SecondoCampo, '' As Data From giocatori " &
					"where month(current_date()) = month(Convert(datadinascita, DateTime)) and day(current_date()) = day(convert(datadinascita, DateTime))"
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Ritorno = ""
					Do Until Rec.Eof()
						Dim SecondoCampo As String = ""

						If Rec("Cosa").Value = "Compleanno" Then
							If Rec("SecondoCampo").Value <> "" Then
								Dim Categorie() As String = Rec("SecondoCampo").Value.ToString.Split("-")

								For Each c As String In Categorie
									Sql = "Select * From Categorie Where idCategoria=" & c
									Rec2 = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
									If Rec2.Eof = True Then
										SecondoCampo = ""
									Else
										SecondoCampo = Rec2("Descrizione").Value & " / "
									End If
									Rec2.Close
								Next

								If SecondoCampo <> "" Then
									SecondoCampo = "(" & SecondoCampo.Substring(1, SecondoCampo.Length - 3) & ")"
								End If
							End If
						Else
							SecondoCampo = Rec("SecondoCampo").Value
						End If
						Ritorno &= Rec("Cosa").Value & ";" & Rec("Id").Value & ";" & Rec("PrimoCampo").Value & ";" & SecondoCampo & ";" & ConverteData(Rec("Data").Value) & "§"

						Rec.MoveNext()
					Loop
					Rec.Close()
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaProssimiEventiUtenti(Squadra As String, Limite As String, Utente As String) As String
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
				Dim Rec2 As Object
				Dim Sql As String
				Dim Altro As String = ""

				Dim Categorie As String = RitornaCategorieUtente(Server.MapPath("."), Conn, Connessione, Utente)

				If Categorie <> "" Then
					Dim Categorie1 As String = ""
					Dim Categorie2 As String = ""

					If Categorie <> "-1" Then
						For Each c As String In Categorie.Split(";")
							If TipoDB = "SQLSERVER" Then
								Categorie1 &= "CHARINDEX('" & c & "-', A.Categorie) > 1 Or "
							Else
								Categorie1 &= "Instr(A.Categorie, '" & c & "-') > 1 Or "
							End If
							Categorie2 &= c & ","
						Next

						Categorie1 = " And (" & Mid(Categorie1, 1, Categorie1.Length - 4) & ")"
						Categorie2 = " Where A.idCategoria In (" & Mid(Categorie2, 1, Categorie2.Length - 1) & ")"
					End If

					If Limite <> "" Then
						Altro = "Top " & Limite
					End If

					Sql = "Select " & Altro & " * From ( " &
						"Select 'Cert. Scad.' As Cosa, A.idGiocatore As Id, 'Certificato medico scaduto' As PrimoCampo, A.Cognome + ' ' + A.Nome As SecondoCampo, B.ScadenzaCertificatoMedico As Data From Giocatori A " &
						"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						"Where A.Eliminato='N' And B.ScadenzaCertificatoMedico Is Not Null And B.ScadenzaCertificatoMedico <> '' " & Categorie1 & " " &
						"And " & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico, 121)", "Convert(B.ScadenzaCertificatoMedico, DateTime)") & " < CURRENT_TIMESTAMP " &
						"Union All " &
						"Select 'Cert. Med.' As Cosa, A.idGiocatore As Id, A.Cognome As PrimoCampo, A.Nome As SecondoCampo, " & IIf(TipoDB = "SQLSERVER", "CONVERT(date, B.ScadenzaCertificatoMedico)", "CONVERT(B.ScadenzaCertificatoMedico, date)") & " As Data From Giocatori A " &
						"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						"Where A.Eliminato='N' And CertificatoMedico = 'S' And A.Eliminato = 'N' And " & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico ,121)", "Convert(B.ScadenzaCertificatoMedico ,DateTime)") & " <= " & IIf(TipoDB = "SQLSERVER", "DateAdd(Day, 30, CURRENT_TIMESTAMP)", "ADDDATE(CURRENT_TIMESTAMP, 30)") & " " & Categorie1 & " " &
						"Union All " &
						"Select 'Partita' As Cosa, idPartita As Id, B.Descrizione As PrimoCampo, C.Descrizione As SecondoCampo, " & IIf(TipoDB = "SQLSERVER", "CONVERT(date, DataOra)", "CONVERT(DataOra, date)") & " As Data From Partite A " &
						"Left Join Categorie B On A.idCategoria = B.idCategoria " &
						"Left Join SquadreAvversarie C On A.idAvversario = C.idAvversario " & Categorie2 & " " &
						"Union All " &
						"Select 'Evento' As Cosa, idEvento As Id, Titolo As PrimoCampo, '' As SecondoCampo, " & IIf(TipoDB = "SQLSERVER", "CONVERT(date, Inizio)", "CONVERT(Inizio, date)") & " As Data From EventiConvocazioni " &
						"Where idTipologia = 2" &
						") A  " &
						"Where Data > " & IIf(TipoDB = "SQLSERVER", "GETDATE()", "CUURENT_DATE()") & " Or Cosa = 'Cert. Scad.' " &
						"Order By Data"
					Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Ritorno = ""
						Do Until Rec.Eof()
							Ritorno &= Rec("Cosa").Value & ";" & Rec("Id").Value & ";" & Rec("PrimoCampo").Value & ";" & Rec("SecondoCampo").Value & ";" & Rec("Data").Value & "§"

							Rec.MoveNext()
						Loop
						Rec.Close()

						If Ritorno = "" Then
							Ritorno = "ERROR: Nessun evento rilevato"
						End If
					End If
				Else
					Ritorno = "ERROR: Nessuna categoria rilevata"
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

		' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Partenza")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object
				Dim Rec2 As Object

				Sql = "Delete From WidgetIndicatori"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Eliminati dati da tabella")

				Dim SenzaQuota As Integer = 0
				Dim CertificatoScadutoAssente As Integer = 0
				Dim SenzaFirma As Integer = 0
				Dim KitNonConsegnato As Integer = 0

				' Giocatori senza quota
				Sql = "Select " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " " &
					"From " &
					"Giocatori A " &
					"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
					"Left Join Quote C On B.idQuota = C.idQuota And C.Eliminato = 'N' " &
					"Where A.Eliminato = 'N' And C.Descrizione Is Null"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Problema lettura recordset senza quota: " & Ritorno)
					Return Ritorno
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	SenzaQuota = 0
					'Else
					SenzaQuota = "" & Rec(0).Value
					'End If
					Rec.Close()
				End If
				' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Senza quota: " & SenzaQuota)

				' Certificato scaduto / assente
				Sql = "Select " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " From Giocatori A " &
						"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						"Where A.Eliminato = 'N' And "
				Sql &= " ((B.ScadenzaCertificatoMedico Is Not Null And B.ScadenzaCertificatoMedico <> '' And " & IIf(TipoDB = "SQLSERVER", "Convert(DateTime, B.ScadenzaCertificatoMedico ,121)", "Convert(B.ScadenzaCertificatoMedico ,DateTime)") & " <= CURRENT_TIMESTAMP And B.CertificatoMedico = 'S') Or "
				Sql &= " (B.CertificatoMedico Is Null Or B.CertificatoMedico = '' Or B.CertificatoMedico = 'N' Or (B.CertificatoMedico = 'S' And (B.ScadenzaCertificatoMedico Is Null Or B.ScadenzaCertificatoMedico = ''))))"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Problema lettura recordset certificato scaduto / assente: " & Ritorno)
					Return Ritorno
				Else
					'If Rec(0).Value Is DBNull.Value Then
					'	CertificatoScadutoAssente = 0
					'Else
					CertificatoScadutoAssente = "" & Rec(0).Value
					'End If
					Rec.Close()
				End If
				' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Certificati scaduti / assenti: " & CertificatoScadutoAssente)

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
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Problema lettura recordset anni: " & Ritorno)
					Return Ritorno
				Else
					If Rec.Eof() = False Then
						IscrFirmaEntrambi = "" & Rec("iscrFirmaEntrambi").Value
					End If
					Rec.Close
				End If
				' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Iscrizione firma entrambi: " & IscrFirmaEntrambi)

				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & idSquadra
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Problema lettura recordset dati squadra: " & Ritorno)
					Return Ritorno
				Else
					If Rec.Eof() = False Then
						NomeSquadra = "" & Rec("Descrizione").Value
					End If
					Rec.Close
				End If
				' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Nome Squadra: " & NomeSquadra)

				Dim wsImm As New wsImmagini

				Sql = "Select * From Giocatori A " &
						"Left Join GiocatoriDettaglio B On A.idGiocatore = B.idGiocatore " &
						"Where A.Eliminato = 'N' "
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Problema lettura recordset dettaglio giocatori: " & Ritorno)
					Return Ritorno
				Else
					Do Until Rec.Eof()
						Dim urlFirma As String = ""
						Dim CiSonoFirme As Boolean = True

						' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Giocatore " & Rec("idGiocatore").Value & ": " & Rec("Cognome").Value & " " & Rec("Nome").Value)

						If "" & Rec("Maggiorenne").Value = "S" Then
							' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Maggiorenne")
							If "" & Rec("AbilitaFirmaGenitore3").Value = "S" Then
								' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma abilitata maggiorenne")
								urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_3.kgb"
								If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "3_1").Contains(StringaErrore) Then
									CiSonoFirme = False
								Else
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente maggiorenne")
									CiSonoFirme = True
								End If
							Else
								If "" & Rec("FirmaAnalogicaGenitore3").Value = "N" Then
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma NON abilitata maggiorenne")
									CiSonoFirme = False
								End If
							End If
						Else
							If "" & Rec("GenitoriSeparati").Value = "S" Then
								' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Genitori separati")
								If "" & Rec("AffidamentoCongiunto").Value = "S" Then
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Affidamento congiunto")
									If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
										' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica abilitata padre")
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
										If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "1_1").Contains(StringaErrore) Then
											CiSonoFirme = False
										Else
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente padre")
											CiSonoFirme = True
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica padre")
											CiSonoFirme = False
										End If
									End If

									If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
										' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica abilitata madre")
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
										If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "2_1").Contains(StringaErrore) Then
											CiSonoFirme = False
										Else
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente madre")
											CiSonoFirme = True
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica madre")
											CiSonoFirme = False
										End If
									End If
								Else
									If "" & Rec("idTutore").Value = "1" Then
										' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Tutore padre")
										If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica abilitata padre")
											urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
											If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "1_1").Contains(StringaErrore) Then
												CiSonoFirme = False
											Else
												' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente padre")
												CiSonoFirme = True
											End If
										Else
											If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
												' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica padre")
												CiSonoFirme = False
											End If
										End If
									Else
										' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Tutore madre")
										If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica abilitata madre")
											urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
											If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "2_1").Contains(StringaErrore) Then
												CiSonoFirme = False
											Else
												' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente madre")
												CiSonoFirme = True
											End If
										Else
											If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
												' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica madre")
												CiSonoFirme = False
											End If
										End If
									End If
								End If
							Else
								If IscrFirmaEntrambi = "S" Then
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Iscrizione firma entrambi")
									If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
										' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica padre")
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
										If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "1_1").Contains(StringaErrore) Then
											CiSonoFirme = False
										Else
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente padre")
											CiSonoFirme = True
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica padre")
											CiSonoFirme = False
										End If
									End If

									If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
										' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica madre")
										urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
										If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "2_1").Contains(StringaErrore) Then
											CiSonoFirme = False
										Else
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente madre")
											CiSonoFirme = True
										End If
									Else
										If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica madre")
											CiSonoFirme = False
										End If
									End If
								Else
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Iscrizione singolo genitore")
									If "" & Rec("Genitore1").Value <> "" Then
										' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Padre presente")
										If "" & Rec("AbilitaFirmaGenitore1").Value = "S" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica padre")
											urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_1.kgb"
											If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "1_1").Contains(StringaErrore) Then
												CiSonoFirme = False
											Else
												' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente padre")
												CiSonoFirme = True
											End If
										Else
											If "" & Rec("FirmaAnalogicaGenitore1").Value = "N" Then
												' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica padre")
												CiSonoFirme = False
											End If
										End If
									Else
										If "" & Rec("Genitore2").Value <> "" Then
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Madre presente")
											If "" & Rec("AbilitaFirmaGenitore2").Value = "S" Then
												' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma elettronica madre")
												urlFirma = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & Rec("idGiocatore").Value & "_2.kgb"
												If wsImm.RitornaImmagineDB(Squadra, "Firme", Rec("idGiocatore").Value, "2_1").Contains(StringaErrore) Then
													CiSonoFirme = False
												Else
													' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma presente madre")
													CiSonoFirme = True
												End If
											Else
												If "" & Rec("FirmaAnalogicaGenitore2").Value = "N" Then
													' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Firma analogica madre")
													CiSonoFirme = False
												Else
													CiSonoFirme = False
												End If
											End If
										Else
											' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Padre e Madre assenti")
											CiSonoFirme = False
										End If
									End If

								End If
							End If
						End If

						If CiSonoFirme = False Then
							' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Incremento il senza firma: " & SenzaFirma)
							SenzaFirma += 1
						End If

						' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "------------------------------------------------------")

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
						Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
						If TypeOf (Rec2) Is String Then
							Ritorno = Rec2
							' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Problema lettura recordset taglie: " & Ritorno)
							Return Ritorno
						Else
							Dim Tutto As Boolean = True
							Dim Qualcosa As Boolean = False

							If Rec2.Eof() Then
								' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Recordset kit: Nessun valore")
								Tutto = False
							Else
								Do Until Rec2.Eof()
									If NomeKit = "" Then
										NomeKit = "" & Rec2("NomeKit").Value
										TagliaKit = "" & Rec2("Taglia").Value
									End If
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Nome Kit: " & NomeKit)
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Taglia Kit: " & TagliaKit)

									'If Val("" & Rec2("QuantitaConsegnata").Value) > 0 Then
									Qualcosa = True
									If Val("" & Rec2("QuantitaConsegnata").Value) < Val("" & Rec2("Quantita").Value) Then
										Tutto = False
									End If
									' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Quantità consegnata " & Val("" & Rec2("QuantitaConsegnata").Value) & " Totale " & Val("" & Rec2("Quantita").Value) & " : " & Tutto)
									'Else
									'	Tutto = False
									'End If

									Rec2.MoveNext()
								Loop
								Rec2.Close()
							End If

							' Kit consegnato No
							Dim Preso As Boolean = False

							If Tutto = False Then
								KitNonConsegnato += 1
								' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "Kit non completo")
							End If
						End If

						' ScriveLog(Server.MapPath("."), Squadra, "RitornaIndicatori", "------------------------------------------------------")

						Rec.MoveNext()
					Loop
				End If
				Rec.Close()

				Sql = "Insert Into WidgetIndicatori Values (" & SenzaQuota & ", " & CertificatoScadutoAssente & ", " & SenzaFirma & ", " & KitNonConsegnato & ")"
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				Ritorno = SenzaQuota & ";" & CertificatoScadutoAssente & ";" & SenzaFirma & ";" & KitNonConsegnato & ";"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaIndicatori(Squadra As String, Refresh As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			If Refresh = "" Or Refresh = "N" Then
				Dim Conn As Object = New clsGestioneDB(Squadra)
				Dim Trovato As Boolean = False

				If TypeOf (Conn) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
				Else
					Dim Rec As Object
					Dim Sql As String = "Select * From WidgetIndicatori"
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof() Then
							Trovato = True
							Ritorno &= Rec("SenzaQuota").Value & ";" & Rec("CertificatoScadutoAssente").Value & ";" & Rec("SenzaFirma").Value & ";" & Rec("KitNonConsegnato").Value
							Rec.Close()
						End If
					End If
				End If

				If Not Trovato Then
					Ritorno = CreaIndicatori(Squadra)
				End If
			Else
				Ritorno = CreaIndicatori(Squadra)
			End If
		End If

		Return Ritorno
	End Function
End Class