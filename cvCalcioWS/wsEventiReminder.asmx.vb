Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://cvcalcio_evrem.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsEventiReminder
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaNuovoID(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idEvento As String = "-1"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				'Dim idUtente As String = ""

				Sql = "SELECT Max(idEvento)+1 FROM EventiReminder"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
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
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				'Dim idUtente As String = ""

				Sql = "SELECT idEvento, idTipologia, Titolo collate Latin1_General_CI_AS As Titolo, Inizio collate Latin1_General_CI_AS As Inizio, " &
					"Fine collate Latin1_General_CI_AS As Fine,  " &
					"TuttiIGiorni collate Latin1_General_CI_AS As TuttiIGiorni, ColorePrimario collate Latin1_General_CI_AS As ColorePrimario, ColoreSecondario collate Latin1_General_CI_AS As ColoreSecondario,  " &
					"metaLocation collate Latin1_General_CI_AS As metaLocation, metaNotes collate Latin1_General_CI_AS As metaNotes, idPartita " &
					"FROM [dbo].[EventiReminder] " &
					"Union All  " &
					"SELECT idEvento, idTipologia, Titolo collate Latin1_General_CI_AS As Titolo, Inizio collate Latin1_General_CI_AS As Inizio,  " &
					"Fine collate Latin1_General_CI_AS As Fine,  " &
					"TuttiIGiorni collate Latin1_General_CI_AS As TuttiIGiorni, ColorePrimario collate Latin1_General_CI_AS As ColorePrimario,  " &
					"ColoreSecondario collate Latin1_General_CI_AS As ColoreSecondario,  " &
					"metaLocation collate Latin1_General_CI_AS As metaLocation, metaNotes collate Latin1_General_CI_AS As metaNotes, idPartita " &
					"FROM [dbo].[EventiConvocazioni]"
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof()
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

						Rec.MoveNext()
					Loop
					Rec.Close()
				End If
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
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Insert Into EventiReminder Values (" &
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
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Rec As Object
				Dim ritEliminazione As String = ""

				Sql = "Select * From EventiReminder Where idEvento = " & idEvento
				Rec = Conn.LeggeQuery(Server.MapPath("."),   Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof() Then
						If Val(Rec("idTipologia").Value) = 1 Then
							Dim idPartita As String = Rec("idPartita").Value

							If idPartita Is DBNull.Value Then
								ritEliminazione = "*"
							Else
								ritEliminazione = EliminaPartita(Server.MapPath("."), Squadra, idAnno, idPartita)
							End If
						End If
					End If
					Rec.Close()
				End If

				If ritEliminazione = "*" Then
					Sql = "Delete From EventiReminder Where idEvento = " & idEvento
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = new clsGestioneDB

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""

				Sql = "Update EventiReminder Set " &
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
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function
End Class