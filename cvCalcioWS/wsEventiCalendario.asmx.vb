Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://cvcalcio_evcal.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsEventiCalendario
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaNuovoID(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idEvento As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				'Dim idUtente As String = ""

				Sql = "SELECT Max(idEvento)+1 FROM EventiCalendario"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec(0).Value Is DBNull.Value Then
						idEvento = "1"
					Else
						idEvento = Rec(0).Value.ToString
					End If
				End If
				Rec.Close()
			End If
		End If

		Return idEvento
	End Function

	<WebMethod()>
	Public Function RitornaEventi(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idArbitro As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				'Dim idUtente As String = ""

				Sql = "SELECT * From EventiCalendario"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("idEvento").Value & ";"
						Ritorno &= Rec("idTipologia").Value & ";"
						Ritorno &= Rec("Titolo").Value & ";"
						Ritorno &= Rec("Inizio").Value & ";"
						Ritorno &= Rec("Fine").Value & ";"
						Ritorno &= Rec("TuttiIGiorni").Value & ";"
						Ritorno &= Rec("ColorePrimario").Value & ";"
						Ritorno &= Rec("ColoreSecondario").Value & ";"
						Ritorno &= Rec("metaLocation").Value & ";"
						Ritorno &= Rec("metaNotes").Value & ";"
						Ritorno &= Rec("idPartita").Value & ";"
						Ritorno &= "§"

						Rec.MoveNext
					Loop
				End If
				Rec.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaEvento(Squadra As String, idEvento As String, idTipologia As String, Titolo As String, Inizio As String, Fine As String, TuttiIGiorni As String,
								ColorePrimario As String, ColoreSecondario As String, metaLocation As String, metaNotes As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idArbitro As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Insert Into EventiCalendario Values (" &
					" " & idEvento & ", " &
					" " & idTipologia & ", " &
					"'" & Titolo.Replace("'", "''").Replace(";", ",") & "', " &
					"'" & Inizio & "', " &
					"'" & Fine & "', " &
					"'" & TuttiIGiorni & "', " &
					"'" & ColorePrimario & "', " &
					"'" & ColoreSecondario & "', " &
					"'" & metaLocation.Replace("'", "''").Replace(";", "-") & "', " &
					"'" & metaNotes.Replace("'", "''").Replace(";", "-") & "', " &
					"'" & idPartita & "' " &
					")"
				Ritorno = EsegueSql(Conn, Sql, Connessione)
				'If Ritorno <> "*" Then
				'	Ritorno = Sql
				'End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaEvento(Squadra As String, idAnno As String, idEvento As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idArbitro As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim ritEliminazione As String = ""

				Sql = "Select * From EventiCalendario Where idEvento = " & idEvento
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						If Val(Rec("idTipologia").Value) = 1 Then
							Dim idPartita As String = Rec("idPartita").Value

							If idPartita Is DBNull.Value Then
								ritEliminazione = "*"
							Else
								ritEliminazione = EliminaPartita(Squadra, idAnno, idPartita)
							End If
						End If
					End If
					Rec.Close
				End If

				If ritEliminazione = "*" Then
					Sql = "Delete From EventiCalendario Where idEvento = " & idEvento
					Ritorno = EsegueSql(Conn, Sql, Connessione)
				Else
					Ritorno = ritEliminazione
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaEvento(Squadra As String, idEvento As String, idTipologia As String, Titolo As String, Inizio As String, Fine As String, TuttiIGiorni As String,
								ColorePrimario As String, ColoreSecondario As String, metaLocation As String, metaNotes As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idArbitro As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Update EventiCalendario Set " &
					"idTipologia = " & idTipologia & ", " &
					"Titolo = '" & Titolo.Replace("'", "''").Replace(";", ",") & "', " &
					"Inizio = '" & Inizio & "', " &
					"Fine = '" & Fine & "', " &
					"TuttiIGiorni = '" & TuttiIGiorni & "', " &
					"ColorePrimario = '" & ColorePrimario & "', " &
					"ColoreSecondario = '" & ColoreSecondario & "', " &
					"metaLocation = '" & metaLocation.Replace("'", "''").Replace(";", ",") & "', " &
					"metaNotes = '" & metaNotes.Replace("'", "''").Replace(";", ",") & "', " &
					"idPartita = '" & idPartita & "' " &
					"Where idEvento = " & idEvento
				Ritorno = EsegueSql(Conn, Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function
End Class