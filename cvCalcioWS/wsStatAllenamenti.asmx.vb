Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_stat_allti.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsStatAllenamenti
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaStatAllenamentiCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String, Mese As String) As String
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
				Dim sMese As String = "/"

				Select Case Mese
					Case "Gennaio"
						sMese = "/01/"
					Case "Febbraio"
						sMese = "/02/"
					Case "Marzo"
						sMese = "/03/"
					Case "Aprile"
						sMese = "/04/"
					Case "Maggio"
						sMese = "/05/"
					Case "Giugno"
						sMese = "/06/"
					Case "Luglio"
						sMese = "/07/"
					Case "Agosto"
						sMese = "/08/"
					Case "Settembre"
						sMese = "/09/"
					Case "Ottobre"
						sMese = "/10/"
					Case "Novembre"
						sMese = "/11/"
					Case "Dicembre"
						sMese = "/12/"
				End Select

				Try
					Sql = "Select B.idGiocatore, B.Cognome, B.Nome, B.Descrizione,  B.Presenze, B.Totale, B.Presenze/B.Totale*100 As Perc, B.NumeroMaglia From ( " &
						"Select A.idGiocatore, A.Cognome, A.Nome, A.Descrizione,  A.Presenze, (SELECT Count(*) From Allenamenti " &
						"Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Instr(Datella,'" & sMese & "')>0  And Progressivo=0) As Totale, A.NumeroMaglia From ( " &
						"SELECT Giocatori.idGiocatore, Cognome, Nome, Ruoli.Descrizione,  Count(*) As Presenze, Giocatori.NumeroMaglia " &
						"FROM (Allenamenti LEFT JOIN Giocatori ON (Allenamenti.idAnno = Giocatori.idAnno) AND (Allenamenti.idGiocatore=Giocatori.idGiocatore) AND (Allenamenti.idCategoria = Giocatori.idCategoria)) " &
						"LEFT Join Ruoli On Giocatori.idRuolo=Ruoli.idRuolo " &
						"WHERE Allenamenti.idCategoria=" & idCategoria & " And Allenamenti.idAnno=" & idAnno & " And Giocatori.idGiocatore Is Not Null And Instr(Datella,'" & sMese & "')>0 " &
						"Group By Giocatori.idGiocatore, Cognome, Nome, Ruoli.Descrizione, Giocatori.NumeroMaglia " &
						") A) B " &
						"Order By 2"

					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna statistica di allenamento rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("Descrizione").Value.ToString.Trim & ";" &
									Rec("Presenze").Value.ToString.Trim & ";" &
									Rec("Totale").Value.ToString.Trim & ";" &
									Rec("Perc").Value.ToString.Trim & ";" &
									Rec("NumeroMaglia").Value.ToString.Trim & ";" &
									"§"
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
	Public Function RitornaInfo(Squadra As String, ByVal idAnno As String, idCategoria As String, idGiocatore As String, Mese As String) As String
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
				Dim sMese As String = "/"

				Select Case Mese
					Case "Gennaio"
						sMese = "/01/"
					Case "Febbraio"
						sMese = "/02/"
					Case "Marzo"
						sMese = "/03/"
					Case "Aprile"
						sMese = "/04/"
					Case "Maggio"
						sMese = "/05/"
					Case "Giugno"
						sMese = "/06/"
					Case "Luglio"
						sMese = "/07/"
					Case "Agosto"
						sMese = "/08/"
					Case "Settembre"
						sMese = "/09/"
					Case "Ottobre"
						sMese = "/10/"
					Case "Novembre"
						sMese = "/11/"
					Case "Dicembre"
						sMese = "/12/"
				End Select

				Try
					Sql = "SELECT Allenamenti.Datella, Allenamenti.Orella " &
						"FROM Allenamenti " &
						"WHERE Allenamenti.idAnno=" & idAnno & " AND Allenamenti.idCategoria=" & idCategoria & " AND Allenamenti.idGiocatore=" & idGiocatore & " And Instr(Datella,'" & sMese & "')>0 " &
						"Order By Datella, Orella"

					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna info di allenamento rilevata"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("Datella").Value.ToString & ";" &
									Rec("Orella").Value.ToString.Trim & ";" &
									"§"
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

End Class