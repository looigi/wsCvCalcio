Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://risposte.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsRisposte
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function GeneraRisposta(Squadra As String, Risposta As String, idPartita As String, idGiocatore As String, Tipo As String) As String
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
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Sql = "Select * From ConvocatiPartiteRisposte Where idPartita=" & idPartita & " And idGiocatore=" & idGiocatore
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
					Ok = False
				Else
					If Rec.Eof Then
						Rec.Close()

						Sql = "Insert Into ConvocatiPartiteRisposte Values (" & idPartita & ", " & idGiocatore & ", '" & Mid(Risposta, 1, 1) & "')"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno = "*" Then
							Sql = "Select A.DataOra, Casa, C.Descrizione As Squadra1, D.Descrizione As Squadra2, B.EMail, E.Descrizione As Campo, E.Indirizzo, F.Lat, F.Lon, " &
								"A.DataOraAppuntamento, A.LuogoAppuntamento, A.Mezzotrasporto " &
								"From Partite A " &
								"Left Join Allenatori B On A.idAllenatore = B.idAllenatore " &
								"Left Join Categorie C On A.idCategoria = C.idCategoria " &
								"Left Join SquadreAvversarie D On A.idAvversario = D.idAvversario " &
								"Left Join CampiAvversari E On D.idCampo = E.idCampo " &
								"Left Join AvversariCoord F On F.idAvversario = D.idAvversario " &
								"Where idPartita =" & idPartita
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Ok = False
							Else
								If Not Rec.Eof Then
									Dim EMAilAllenatore As String = Rec("EMail").Value
									Dim Casa As String = "" & Rec("Casa").Value
									Dim DataOra As String = "" & Rec("DataOra").Value
									Dim Campo As String = "" & Rec("Campo").Value
									Dim Indirizzo As String = "" & Rec("Indirizzo").value
									Dim Sq1 As String = "" & Rec("Squadra1").Value
									Dim Sq2 As String = "" & Rec("Squadra2").Value
									Dim LatLon As String = "" & Rec("Lat").Value & "," & Rec("Lon").Value
									Dim DataOraApp As String = "" & Rec("DataOraAppuntamento").Value
									If DataOraApp <> "" And
										DataOraApp = FormatDateTime(DataOra, DateFormat.LongDate) & " " & FormatDateTime(DataOra, DateFormat.ShortTime) Then
									End If
									Dim LuogoApp As String = "" & Rec("LuogoAppuntamento").Value
									Dim MezzoTrasporto As String = "" & Rec("MezzoTrasporto").Value
									Dim Mezzo As String = ""
									If MezzoTrasporto = "P" Then
										Mezzo = "pullman"
									Else
										Mezzo = "auto propria"
									End If
									Dim descPartita As String = ""
									If Casa = "S" Then
										descPartita = Sq1 & "-" & Sq2
									Else
										descPartita = Sq2 & "-" & Sq1
									End If
									Rec.Close()

									Sql = "Select * From Giocatori Where idGiocatore=" & idGiocatore
									Rec = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
										Ok = False
									Else
										If Not Rec.Eof Then
											Dim Cognome As String = "" & Rec("Cognome").Value
											Dim Nome As String = "" & Rec("Nome").Value
											Rec.Close()

											Dim ma As New mail
											Dim Altro As String
											If Risposta = "SI" Then
												Altro = "positiva"
											Else
												Altro = "negativa"
											End If
											Dim Oggetto As String = "Risposta " & Altro & " Giocatore " & Cognome & " " & Nome & " per partita " & descPartita
											Dim Body As String = "Il giocatore " & Cognome & " " & Nome & " ha risposto in maniera " & Altro & " alla convocazione per la partita " & descPartita & " "
											Body &= "che si giocherà " & FormatDateTime(DataOra, DateFormat.LongDate) & " " & FormatDateTime(DataOra, DateFormat.ShortTime)
											If Casa = "S" Then
												Body &= " in casa."
											Else
												Dim url As String = "https://www.google.it/maps/place/" & LatLon
												Body &= " al campo " & Campo & ", indirizzo <a href=""" & url & """>" & Indirizzo & "</a>."
											End If
											Dim url2 As String = "https://www.google.it/maps/place/" & LuogoApp
											Body &= "<br /><br />Appuntamento: " & FormatDateTime(DataOraApp, DateFormat.LongDate) & " " & FormatDateTime(DataOraApp, DateFormat.ShortTime) & " <a href=""" & url2 & """>" & LuogoApp & "</a> tramite " & Mezzo
											Ritorno = ma.SendEmail(Squadra, "", Oggetto, Body, EMAilAllenatore, {""})
										End If
									End If
								Else
									Ritorno = StringaErrore & " Risposta già inviata"
									Ok = False
								End If
							End If
						End If
					End If

					If Ok Then
						Sql = "commit"
						Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
					Else
						Sql = "rollback"
						Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class