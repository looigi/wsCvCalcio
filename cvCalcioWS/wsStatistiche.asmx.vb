Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_stat.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsStatistiche
    Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaStatisticheAvversari(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "SELECT SquadreAvversarie.idAvversario, SquadreAvversarie.Descrizione, Count(*) AS Quante "
				Sql &= "FROM (Partite LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) "
				Sql &= "LEFT JOIN Categorie On Partite.idAnno=Categorie.idAnno And Categorie.idCategoria=Partite.idCategoria "
				If SoloAnno = "S" Then
					Sql &= "WHERE Partite.idAnno=" & idAnno & " And Partite.Giocata='S' And Categorie.idCategoria=" & idCategoria & " "
				Else
					Sql &= "WHERE Partite.Giocata='S' And Categorie.idCategoria=" & idCategoria & " "
				End If
				Sql &= "GROUP BY SquadreAvversarie.idAvversario, SquadreAvversarie.Descrizione "
				Sql &= "ORDER BY 3 DESC"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("idAvversario").Value & ";" & Rec("Descrizione").Value & ";" & Rec("Quante").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheConvocati(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "SELECT Giocatori.idGiocatore, Cognome, Nome, Count(*) As Quanti, NumeroMaglia "
				Sql &= "FROM ((Giocatori INNER JOIN Partite ON Giocatori.idAnno = Partite.idAnno) "
				Sql &= "INNER JOIN Convocati ON Partite.idPartita = Convocati.idPartita And Giocatori.idGiocatore=Convocati.idGiocatore) "
				Sql &= "INNER JOIN Categorie On Giocatori.idCategoria=Categorie.idCategoria And Giocatori.idAnno=Categorie.idAnno "
				If SoloAnno = "S" Then
					Sql &= "WHERE Giocatori.idAnno= " & idAnno & " And Categorie.idCategoria=" & idCategoria & " "
				Else
					Sql &= "WHERE Categorie.idCategoria=" & idCategoria & " "
				End If
				Sql &= "Group By Giocatori.idGiocatore, Cognome, Nome, NumeroMaglia "
				Sql &= "Order By 4 Desc,2,3"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("Quanti").Value & ";" & Rec("NumeroMaglia").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheMarcatori(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ": " & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				Sql = "SELECT q.idGiocatore, q.Cognome, q.Nome, Sum(Goal) AS GoalFinali, q.NumeroMaglia FROM("
				Sql &= "Select Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) AS Goal, Giocatori.NumeroMaglia "
				Sql &= "FROM ((Marcatori LEFT  JOIN Partite ON Marcatori.idPartita = Partite.idPartita) "
				Sql &= "LEFT JOIN Giocatori ON Marcatori.idGiocatore = Giocatori.idGiocatore) "
				Sql &= "LEFT JOIN Categorie On Giocatori.idCategoria=Categorie.idCategoria And Categorie.idAnno=Giocatori.idAnno "
				If SoloAnno = "S" Then
					Sql &= "WHERE Partite.idAnno=" & idAnno & " And Partite.Giocata='S' And Cognome Is Not Null And Categorie.idCategoria=" & idCategoria & " "
				Else
					Sql &= "WHERE Partite.Giocata='S' And Cognome Is Not Null And Categorie.idCategoria=" & idCategoria & " "
				End If
				Sql &= "GROUP BY Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Giocatori.NumeroMaglia "
				Sql &= "Union All "
				Sql &= "Select Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.NumeroMaglia "
				Sql &= "FROM ((RisultatiAggiuntiviMarcatori LEFT JOIN Partite On RisultatiAggiuntiviMarcatori.idPartita = Partite.idPartita) "
				Sql &= "LEFT JOIN Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore) "
				Sql &= "LEFT JOIN Categorie On Giocatori.idCategoria=Categorie.idCategoria And Categorie.idAnno=Giocatori.idAnno "
				If SoloAnno = "S" Then
					Sql &= "WHERE Partite.idAnno=" & idAnno & " And Partite.Giocata='S' And Cognome Is Not Null And Categorie.idCategoria=" & idCategoria & "  And Giocatori.idAnno=Partite.idanno "
				Else
					Sql &= "WHERE Partite.Giocata='S' And Cognome Is Not Null And Categorie.idCategoria=" & idCategoria & "  And Giocatori.idAnno=Partite.idanno "
				End If
				Sql &= "GROUP BY Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Giocatori.NumeroMaglia "
				Sql &= ") AS q "
				Sql &= "Group BY q.idGiocatore, q.Cognome, q.Nome, q.NumeroMaglia "
				Sql &= "ORDER BY 4 DESC, 2, 3"

				Dim Ok As Boolean = True

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("GoalFinali").Value & ";" & Rec("NumeroMaglia").Value & ";GOAL;§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try

				If Ok Then
					Sql = "SELECT RigoriPropri.idGiocatore, Ruoli.Descrizione, Giocatori.Cognome, Giocatori.Nome, Giocatori.NumeroMaglia, Sum(RigoriPropri.Termine) As Rigori " &
						"From ((RigoriPropri Left Join Giocatori On RigoriPropri.idGiocatore=Giocatori.idGiocatore And RigoriPropri.idAnno = Giocatori.idAnno) Left Join Ruoli On Giocatori.idRuolo = Ruoli.idRuolo) " &
						"Where RigoriPropri.idAnno=" & idAnno & " And Termine=1 " &
						"Group By RigoriPropri.idGiocatore, Ruoli.Descrizione, Giocatori.Cognome, Giocatori.Nome, Giocatori.NumeroMaglia"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
					Else
						Do Until Rec.Eof
							' 448;Centrocampista;Cataldi Lorenzo;;;14;-1;
							' RigoriPropri &= Rec2("idGiocatore").Value & "!"
							' RigoriPropri &= Rec2("Descrizione").Value & "!"
							' RigoriPropri &= Rec2("Giocatore").Value & "!"
							' RigoriPropri &= "!"
							' RigoriPropri &= "!"
							' RigoriPropri &= Rec2("NumeroMaglia").Value & "!"
							' RigoriPropri &= Rec2("Termine").Value & "!"
							' RigoriPropri &= "%"
							Ritorno &= Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("Rigori").Value & ";" & Rec("NumeroMaglia").Value & ";RIGORE;§"

							Rec.MoveNext
						Loop
					End If
					Rec.Close
				End If

				Conn.Close()
			End If
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheRisultati(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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

				Sql = "Drop Table Appoggio"
				EsegueSql(Conn, Sql, Connessione)

				Sql = "Select * Into Appoggio From ( "
				'Sql &= "Select 1 As Descrizione, Partite.idPartita As Partita, Partite.Casa, Count(*) As Valore From (Partite Inner Join RisultatiAggiuntiviMarcatori On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) "
				'If SoloAnno = "S" Then
				'    Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata ='S' And Partite.idCategoria=" & idCategoria & " "
				'Else
				'    Sql &= "Where Partite.Giocata ='S' And Partite.idCategoria=" & idCategoria & " "
				'End If
				'Sql &= "Group By Partite.idPartita, Partite.Casa "

				Sql &= "Select Descrizione, Partita,Casa, Sum(Valo) As Valore From ( "
				Sql &= "Select 1 As Descrizione, Partite.idPartita As Partita, Partite.Casa, Count(*) As Valo From (Partite Inner Join RisultatiAggiuntiviMarcatori On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) "
				Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata ='S' And Partite.idCategoria=" & idCategoria & " "
				Sql &= "Group By Partite.idPartita, Partite.Casa "
				Sql &= "Union All "
				Sql &= "Select 1 As Descrizione, Partite.idPartita As Partita, Partite.Casa, 0 As Valo From (Partite left Join RisultatiAggiuntiviMarcatori On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) "
				Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata ='S' And Partite.idCategoria=" & idCategoria & " "
				Sql &= "Group By Partite.idPartita, Partite.Casa "
				Sql &= ") Group By Descrizione, Partita, Casa "

				Sql &= "Union All "
				Sql &= "Select 2 As Descrizione, Partite.idPartita As Partita, Partite.Casa, Sum(IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) As Valore From (Partite Left Join RisultatiAggiuntivi On Partite.idPartita = RisultatiAggiuntivi.idPartita) "
				If SoloAnno = "S" Then
					Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata ='S' And Partite.idCategoria=" & idCategoria & " "
				Else
					Sql &= "Where Partite.Giocata ='S' And Partite.idCategoria=" & idCategoria & " "
				End If
				Sql &= "Group By Partite.idPartita, Partite.Casa)"
				EsegueSql(Conn, Sql, Connessione)

				Try
					Sql = "Select 'Giocate Totali:' As Descrizione, Count(*) As Valore From Partite "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " "
					Else
						Sql &= "Where idCategoria= " & idCategoria & " "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Goal Fatti Totali:' As Descrizione, Count(*) As Valore From (RisultatiAggiuntiviMarcatori Left Join "
					Sql &= "Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata='S' And Partite.idCategoria = " & idCategoria & " "
					Else
						Sql &= "Where Partite.Giocata='S' And Partite.idCategoria = " & idCategoria & " "
					End If
					Sql &= "Union All "
					Sql &= "SELECT 'Goal Subiti Totali:' As Descrizione, Sum(IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) As Valore "
					Sql &= "From (Partite Left Join RisultatiAggiuntivi On Partite.idPartita=RisultatiAggiuntivi.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " "
					Else
						Sql &= "Where idCategoria = " & idCategoria & " "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Giocate Casa:' As Descrizione, Count(*) As Valore From Partite "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " And Casa='S' "
					Else
						Sql &= "Where idCategoria = " & idCategoria & " And Casa='S' "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Goal Fatti In Casa:' As Descrizione, Count(*) As Valore From (RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata='S' And Partite.idCategoria= " & idCategoria & " And Partite.Casa = 'S' "
					Else
						Sql &= "Where Partite.Giocata='S' And Partite.idCategoria= " & idCategoria & " And Partite.Casa = 'S' "
					End If
					Sql &= "Union All "
					Sql &= "SELECT 'Goal Subiti In Casa:' As Descrizione, Sum(IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) As Valore "
					Sql &= "From (Partite Left Join RisultatiAggiuntivi On Partite.idPartita=RisultatiAggiuntivi.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " And Partite.Casa='S' "
					Else
						Sql &= "Where idCategoria = " & idCategoria & " And Partite.Casa='S' "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Giocate Fuori:' As Descrizione, Count(*) As Valore From Partite "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " And Casa='N' "
					Else
						Sql &= "Where idCategoria = " & idCategoria & " And Casa='N' "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Goal Fatti Fuori Casa:' As Descrizione, Count(*) As Valore From (RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata='S' And Partite.idCategoria = " & idCategoria & " And Partite.Casa = 'N' "
					Else
						Sql &= "Where Partite.Giocata='S' And Partite.idCategoria = " & idCategoria & " And Partite.Casa = 'N' "
					End If
					Sql &= "Union All "
					Sql &= "SELECT 'Goal Subiti Fuori Casa:' As Descrizione, Sum(IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) As Valore "
					Sql &= "From (Partite Left Join RisultatiAggiuntivi On Partite.idPartita=RisultatiAggiuntivi.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " And Partite.Casa='N' "
					Else
						Sql &= "Where idCategoria = " & idCategoria & " And Partite.Casa='N' "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Giocate Campo Esterno:' As Descrizione, Count(*) As Valore From Partite "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " And Casa='E' "
					Else
						Sql &= "Where idCategoria = " & idCategoria & " And Casa='E' "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Goal Fatti Campo Esterno:' As Descrizione, Count(*) As Valore From (RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.Giocata='S' And Partite.idCategoria= " & idCategoria & " And Partite.Casa = 'E' "
					Else
						Sql &= "Where Partite.Giocata='S' And Partite.idCategoria = " & idCategoria & " And Partite.Casa = 'E' "
					End If
					Sql &= "Union All "
					Sql &= "Select 'Goal Subiti Campo Esterno:' As Descrizione, Sum(IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) As Valore "
					Sql &= "From (Partite Left Join RisultatiAggiuntivi On Partite.idPartita=RisultatiAggiuntivi.idPartita) "
					If SoloAnno = "S" Then
						Sql &= "Where idAnno=" & idAnno & " And idCategoria = " & idCategoria & " And Partite.Casa='E' "
					Else
						Sql &= "Where idCategoria = " & idCategoria & " And Partite.Casa='E' "
					End If
					Sql &= "Union All "
					Sql &= "Select Iif(Risultato='1','Vittoria',IIf(Risultato='2','Sconfitta','Pareggio')) + ' ' + Iif(Casa='S','Casa',Iif(Casa='N','Fuori','Campo Esterno')) +':' As Descrizione, Count(*) As Valore From ( "
					Sql &= "Select A2.*, iif(Differenza>0,'1',iif(Differenza<0,'2','X')) As Risultato From ( "
					Sql &= "Select A1.*, Propria-Altri As Differenza From ( "
					Sql &= "Select Partita, Casa, (Select Valore From Appoggio Where Descrizione=1 And Partita=A.Partita) As Propria, (Select Valore From Appoggio Where Descrizione=2 And Partita=A.Partita) As Altri "
					Sql &= "From Appoggio As A Where Descrizione=1) As A1) As A2 ) As A3 "
					Sql &= "Group By Risultato, Casa "
					Sql &= "Union All "
					Sql &= "Select  Iif(Risultato='1','Vittoria',IIf(Risultato='2','Sconfitta','Pareggio')) + ' Totali:' As Descrizione, Count(*) As Valore From ( "
					Sql &= "Select A2.*, iif(Differenza>0,'1',iif(Differenza<0,'2','X')) As Risultato From ( "
					Sql &= "Select A1.*, Propria-Altri As Differenza From ( "
					Sql &= "Select Partita, '' As Descrizione, (Select Valore From Appoggio Where Descrizione=1 And Partita=A.Partita) As Propria, (Select Valore From Appoggio Where Descrizione=2 And Partita=A.Partita) As Altri "
					Sql &= "From Appoggio As A Where Descrizione=1) As A1) As A2) As A3 "
					Sql &= "Group By Risultato "
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("Descrizione").Value & ";" & Rec("Valore").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Conn.Close()
			End If
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheMappa(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String
				Dim Ok As Boolean = True
				Dim IndirizzoCasa As String = ""
				Dim Lat As String = ""
				Dim Lon As String = ""

				Sql = "Select * From Anni Where idAnno=" & idAnno
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Ritorno &= "0;0;0;" & Rec("Indirizzo").Value & ";0;" & Rec("Lat").Value & ";" & Rec("Lon").Value & ";0;0;S§"
						IndirizzoCasa = Rec("Indirizzo").Value
						Lat = Rec("Lat").Value
						Lon = Rec("Lon").Value
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
					Ok = False
				End Try
				Rec.Close()

				If Ok Then
					Sql = "SELECT Partite.Casa, Partite.idPartita, SquadreAvversarie.Descrizione As Squadra, CampiAvversari.Descrizione As Campo, CampiAvversari.Indirizzo, "
					Sql &= "Categorie.Descrizione As Categoria, CoordinatePartite.Lat, CoordinatePartite.Lon, CampiEsterni.Descrizione As CampoEsterno, Partite.idAvversario, Partite.DataOra "
					Sql &= "FROM ((((Partite LEFT JOIN CoordinatePartite ON Partite.idPartita = CoordinatePartite.idPartita) "
					Sql &= "LEFT JOIN SquadreAvversarie On Partite.idAvversario=SquadreAvversarie.idAvversario) "
					Sql &= "LEFT JOIN CampiAvversari On SquadreAvversarie.idCampo=CampiAvversari.idCampo) "
					Sql &= "LEFT JOIN Categorie On Categorie.idAnno=Partite.idAnno) "
					Sql &= "LEFT JOIN CampiEsterni On CampiEsterni.idPartita=Partite.idPartita "
					Sql &= "WHERE Partite.idAnno=" & idAnno & " And Categorie.idCategoria=" & idCategoria & " "
					Sql &= "Order By Partite.DataOra"

					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								If Rec("Casa").Value = "N" Then
									Ritorno &= Rec("idPartita").Value & ";" & Rec("Squadra").Value & ";" & Rec("Campo").Value & ";" & Rec("Indirizzo").Value & ";" & Rec("Categoria").Value & ";" & Rec("Lat").Value & ";" & Rec("Lon").Value & ";" & Rec("idAvversario").Value & ";" & Rec("DataOra").Value & ";" & Rec("Casa").Value & "§"
								Else
									If Rec("Casa").Value = "S" Then
										Ritorno &= Rec("idPartita").Value & ";" & Rec("Squadra").Value & ";Casa;" & IndirizzoCasa & ";" & Rec("Categoria").Value & ";" & Lat & ";" & Lon & ";" & Rec("idAvversario").Value & ";" & Rec("DataOra").Value & ";" & Rec("Casa").Value & "§"
									Else
										Ritorno &= Rec("idPartita").Value & ";" & Rec("Squadra").Value & ";" & Rec("CampoEsterno").Value & ";;" & Rec("Categoria").Value & ";" & Rec("Lat").Value & ";" & Rec("Lon").Value & ";" & Rec("idAvversario").Value & ";" & Rec("DataOra").Value & ";" & Rec("Casa").Value & "§"
									End If
								End If

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				Conn.Close()
			End If
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheMinutiGoal(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "SELECT RisultatiAggiuntiviMarcatori.Minuto+1 As Minuto, Count(*) As Quanti "
				Sql &= "FROM RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita "
				Sql &= "Where  Partite.idAnno=" & idAnno & " And Partite.idCategoria=" & idCategoria & " And RisultatiAggiuntiviMarcatori.Minuto Is Not Null "
				Sql &= "Group by RisultatiAggiuntiviMarcatori.Minuto+1 Order By 1"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= "1;" & Rec("Minuto").Value & ";" & Rec("Quanti").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close

						Sql = "Select TempiPrimoTempo + TempiSecondoTempo + TempiTerzoTempo As MinutiSubiti From "
						Sql &= "TempiGoalAvversari Left Join Partite On TempiGoalAvversari.idPartita=Partite.idPartita "
						Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.idCategoria = " & idCategoria & " And (TempiPrimoTempo <> '' Or TempiSecondoTempo <> '' Or TempiTerzoTempo <> '') "

						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim Minuti As String = ""

							Do Until Rec.Eof
								Minuti &= Rec("MinutiSubiti").Value

								Rec.MoveNext
							Loop
							Rec.Close()

							Dim mm() As String = Minuti.Split("#")
							Dim mmm As List(Of String) = New List(Of String)
							Dim nnn As List(Of Integer) = New List(Of Integer)

							For Each m As String In mm
								Dim n As Integer = 0
								Dim Ok As Boolean = False

								For Each m2 As String In mmm
									If m = m2 Then
										nnn.Item(n) += 1
										Ok = True
										Exit For
									End If
									n += 1
								Next

								If Not Ok Then
									mmm.Add(m)
									nnn.Add(1)
								End If
							Next

							For ii As Integer = 0 To mmm.Count - 1
								For kk As Integer = 0 To mmm.Count - 1
									If Val(mmm(ii)) < Val(mmm(kk)) Then
										Dim Appoggio As String = mmm(ii)
										mmm(ii) = mmm(kk)
										mmm(kk) = Appoggio

										Appoggio = nnn(ii)
										nnn(ii) = nnn(kk)
										nnn(kk) = Appoggio
									End If
								Next
							Next

							Dim i As Integer = 0

							For Each min As String In mmm
								If min <> "" Then
									Ritorno &= "2;" & min & ";" & nnn.Item(i) & "§"
								End If
								i += 1
							Next
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Conn.Close()
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheMeteo(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "Select MeteoPartite.Gradi, MeteoPartite.Umidita, MeteoPartite.Pressione "
				Sql &= "From MeteoPartite Left Join Partite On MeteoPartite.idPartita=Partite.idPartita "
				Sql &= "Where Partite.idAnno = " & idAnno & " And Partite.idCategoria = " & idCategoria & " "
				Sql &= "Order By MeteoPartite.idPartita"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= "1;" & Rec("Gradi").Value & ";" & Rec("Umidita").Value & ";" & Rec("Pressione").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()

						Sql = "SELECT MeteoPartite.Tempo, Count(*) As Quanti "
						Sql &= "From MeteoPartite Left Join Partite On MeteoPartite.idPartita=Partite.idPartita "
						Sql &= "Where Partite.idAnno=" & idAnno & " And idCategoria=" & idCategoria & " "
						Sql &= "Group By Tempo "

						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Ritorno &= "2;" & Rec("Tempo").Value & ";" & Rec("Quanti").Value & "§"

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Conn.Close()
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheGoalSegnatiSubiti(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "Select Partite.idPartita, Count(*) As Goals From "
				Sql &= "RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita=RisultatiAggiuntiviMarcatori.idPartita "
				Sql &= "Where Partite.idAnno=" & idAnno & " And idCategoria=" & idCategoria & " "
				Sql &= "Group By Partite.idPartita "
				Sql &= "Order By Partite.idPartita"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= "1;" & Rec("idPartita").Value & ";" & Rec("Goals").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()

						Sql = "Select Partite.idPartita, Partite.Casa, IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0) As Risultato "
						Sql &= "From RisultatiAggiuntivi Left Join Partite On Partite.idPartita=RisultatiAggiuntivi.idPartita "
						Sql &= "Where Partite.idAnno=" & idAnno & " And idCategoria=" & idCategoria & " "
						Sql &= "Order By Partite.idPartita"

						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Dim g() As String = Rec("Risultato").Value.ToString.Split("-")

								If Rec("Casa").Value = "S" Then
									Ritorno &= "2;" & Rec("idPartita").Value & ";" & g(1) & "§"
								Else
									Ritorno &= "2;" & Rec("idPartita").Value & ";" & g(0) & "§"
								End If

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Conn.Close()
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaAndamento(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "Select idPartita, Casa, Sum(GoalPropri) As Goal1, Sum(GoalAvversari) As Goal2 From ("
				Sql &= "Select Partite.idPartita, Partite.Casa, 0 As GoalPropri, Sum(IIf(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0) + IIf(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0) + IIf(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) As GoalAvversari "
				Sql &= "From (RisultatiAggiuntivi Left Join Partite On Partite.idPartita=RisultatiAggiuntivi.idPartita) "
				Sql &= "Where Partite.idAnno=" & idAnno & " And idCategoria=" & idCategoria & " "
				Sql &= "Group By Partite.idPartita, Partite.Casa "
				Sql &= "Union All "
				Sql &= "Select Partite.idPartita, Partite.Casa, Count(*) As GoalPropri, 0 As GoalAvversari "
				Sql &= "From RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita=RisultatiAggiuntiviMarcatori.idPartita "
				Sql &= "Where Partite.idAnno=" & idAnno & " And idCategoria=" & idCategoria & " "
				Sql &= "Group By Partite.idPartita, Partite.Casa "
				Sql &= ") As A Group By idPartita, Casa Order By idPartita"

				Dim Punti As Integer = 0
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Ritorno &= "0;0§"

						Do Until Rec.Eof
							Dim g1 As Integer = Rec("Goal1").Value
							Dim g2 As Integer = Rec("Goal2").Value

							If Rec("Casa").Value = "S" Then
								If g1 > g2 Then
									Punti += 3
								Else
									If g1 = g2 Then
										Punti += 1
									Else
										' Punti -= 3
									End If
								End If
							Else
								If g1 > g2 Then
									Punti += 3
								Else
									If g1 = g2 Then
										Punti += 1
									Else
										' Punti -= 3
									End If
								End If
							End If

							Ritorno &= Rec("idPartita").Value & ";" & Punti & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Conn.Close()
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaTipologiePartite(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "SELECT TipologiePartite.Descrizione, Count(*) As Volte "
				Sql &= "FROM Partite Left Join TipologiePartite On Partite.idTipologia=TipologiePartite.idTipologia "
				Sql &= "Where idAnno=" & idAnno & " And idCategoria= " & idCategoria & " "
				Sql &= "Group By TipologiePartite.Descrizione"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("Descrizione").Value & ";" & Rec("Volte").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Conn.Close()
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaPartiteCasaFuori(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "SELECT iif(Partite.Casa='S','In casa', iif(Partite.Casa='E','Campo esterno','Fuori casa')) As Dove, Count(*) As Volte "
				Sql &= "FROM Partite "
				Sql &= "Where idAnno=" & idAnno & " And idCategoria= " & idCategoria & " "
				Sql &= "Group By  iif(Partite.Casa='S','In casa', iif(Partite.Casa='E','Campo esterno','Fuori casa'))"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("Dove").Value & ";" & Rec("Volte").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Conn.Close()
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaPartiteAllenatore(Squadra As String, idAnno As String, SoloAnno As String, idCategoria As String) As String
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
				Dim Sql As String

				Sql = "SELECT Allenatori.Cognome, Allenatori.Nome, Count(*) As Volte "
				Sql &= "FROM Partite LEFT JOIN Allenatori ON (Partite.idAnno = Allenatori.idAnno) AND (Partite.idAllenatore = Allenatori.idAllenatore) "
				Sql &= "Where Partite.idAnno=" & idAnno & " And Partite.idCategoria= " & idCategoria & " "
				Sql &= "Group By Allenatori.Cognome, Allenatori.Nome"

				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("Cognome").Value & " " & Rec("Nome").Value & ";" & Rec("Volte").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Conn.Close()
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatisticheStagione(Squadra As String, idAnno As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory
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

				Dim PartiteCampionato As New List(Of Integer)
				Dim PartiteCampionatoDove As New List(Of String)
				Dim PartiteCampionatoIN As String = ""
				Dim PartiteAmichevoli As New List(Of Integer)
				Dim PartiteAmichevoliDove As New List(Of String)
				Dim PartiteAmichevoliIN As String = ""
				Dim PartiteTornei As New List(Of Integer)
				Dim PartiteTorneiDove As New List(Of String)
				Dim PartiteTorneiIN As String = ""

				Dim GiocateCampionatoCasa As Integer = 0
				Dim GiocateCampionatoFuori As Integer = 0
				Dim GiocateCampionatoCampoEsterno As Integer = 0

				Dim GiocateAmichevoliCasa As Integer = 0
				Dim GiocateAmichevoliFuori As Integer = 0
				Dim GiocateAmichevoliCampoEsterno As Integer = 0

				Dim GiocateTorneiCasa As Integer = 0
				Dim GiocateTorneiFuori As Integer = 0
				Dim GiocateTorneiCampoEsterno As Integer = 0

				Dim GoalCampionatoCasa As Integer = 0
				Dim GoalCampionatoFuori As Integer = 0
				Dim GoalCampionatoCampoEsterno As Integer = 0

				Dim GoalAmichevoliCasa As Integer = 0
				Dim GoalAmichevoliFuori As Integer = 0
				Dim GoalAmichevoliCampoEsterno As Integer = 0

				Dim GoalTorneiCasa As Integer = 0
				Dim GoalTorneiFuori As Integer = 0
				Dim GoalTorneiCampoEsterno As Integer = 0

				Dim NomiMarcatoriGeneraliCasa As New List(Of String)
				Dim NomiMarcatoriGeneraliFuori As New List(Of String)
				Dim NomiMarcatoriGeneraliCampoEsterno As New List(Of String)

				Dim NomiMarcatoriCampionatoCasa As New List(Of String)
				Dim NomiMarcatoriCampionatoFuori As New List(Of String)
				Dim NomiMarcatoriCampionatoCampoEsterno As New List(Of String)
				Dim MarcatoriCampionatoCasa As Integer = 0
				Dim MarcatoriCampionatoFuori As Integer = 0
				Dim MarcatoriCampionatoCampoEsterno As Integer = 0

				Dim NomiMarcatoriAmichevoliCasa As New List(Of String)
				Dim NomiMarcatoriAmichevoliFuori As New List(Of String)
				Dim NomiMarcatoriAmichevoliCampoEsterno As New List(Of String)
				Dim MarcatoriAmichevoliCasa As Integer = 0
				Dim MarcatoriAmichevoliFuori As Integer = 0
				Dim MarcatoriAmichevoliCampoEsterno As Integer = 0

				Dim NomiMarcatoriTorneiCasa As New List(Of String)
				Dim NomiMarcatoriTorneiFuori As New List(Of String)
				Dim NomiMarcatoriTorneiCampoEsterno As New List(Of String)
				Dim MarcatoriTorneiCasa As Integer = 0
				Dim MarcatoriTorneiFuori As Integer = 0
				Dim MarcatoriTorneiCampoEsterno As Integer = 0

				Dim GoalAvvCampionatoCasa1Tempo As Integer = 0
				Dim GoalAvvCampionatoCasa2Tempo As Integer = 0
				Dim GoalAvvCampionatoCasa3Tempo As Integer = 0
				Dim GoalAvvCampionatoFuori1Tempo As Integer = 0
				Dim GoalAvvCampionatoFuori2Tempo As Integer = 0
				Dim GoalAvvCampionatoFuori3Tempo As Integer = 0
				Dim GoalAvvCampionatoCampoEsterno1Tempo As Integer = 0
				Dim GoalAvvCampionatoCampoEsterno2Tempo As Integer = 0
				Dim GoalAvvCampionatoCampoEsterno3Tempo As Integer = 0

				Dim GoalAvvAmichevoliCasa1Tempo As Integer = 0
				Dim GoalAvvAmichevoliCasa2Tempo As Integer = 0
				Dim GoalAvvAmichevoliCasa3Tempo As Integer = 0
				Dim GoalAvvAmichevoliFuori1Tempo As Integer = 0
				Dim GoalAvvAmichevoliFuori2Tempo As Integer = 0
				Dim GoalAvvAmichevoliFuori3Tempo As Integer = 0
				Dim GoalAvvAmichevoliCampoEsterno1Tempo As Integer = 0
				Dim GoalAvvAmichevoliCampoEsterno2Tempo As Integer = 0
				Dim GoalAvvAmichevoliCampoEsterno3Tempo As Integer = 0

				Dim GoalAvvTorneiCasa1Tempo As Integer = 0
				Dim GoalAvvTorneiCasa2Tempo As Integer = 0
				Dim GoalAvvTorneiCasa3Tempo As Integer = 0
				Dim GoalAvvTorneiFuori1Tempo As Integer = 0
				Dim GoalAvvTorneiFuori2Tempo As Integer = 0
				Dim GoalAvvTorneiFuori3Tempo As Integer = 0
				Dim GoalAvvTorneiCampoEsterno1Tempo As Integer = 0
				Dim GoalAvvTorneiCampoEsterno2Tempo As Integer = 0
				Dim GoalAvvTorneiCampoEsterno3Tempo As Integer = 0

				Dim VittorieCampionatoCasa As Integer = 0
				Dim PareggiCampionatoCasa As Integer = 0
				Dim SconfitteCampionatoCasa As Integer = 0
				Dim VittorieCampionatoFuori As Integer = 0
				Dim PareggiCampionatoFuori As Integer = 0
				Dim SconfitteCampionatoFuori As Integer = 0
				Dim VittorieCampionatoCampoEsterno As Integer = 0
				Dim PareggiCampionatoCampoEsterno As Integer = 0
				Dim SconfitteCampionatoCampoEsterno As Integer = 0

				Dim VittorieAmichevoliCasa As Integer = 0
				Dim PareggiAmichevoliCasa As Integer = 0
				Dim SconfitteAmichevoliCasa As Integer = 0
				Dim VittorieAmichevoliFuori As Integer = 0
				Dim PareggiAmichevoliFuori As Integer = 0
				Dim SconfitteAmichevoliFuori As Integer = 0
				Dim VittorieAmichevoliCampoEsterno As Integer = 0
				Dim PareggiAmichevoliCampoEsterno As Integer = 0
				Dim SconfitteAmichevoliCampoEsterno As Integer = 0

				Dim VittorieTorneiCasa As Integer = 0
				Dim PareggiTorneiCasa As Integer = 0
				Dim SconfitteTorneiCasa As Integer = 0
				Dim VittorieTorneiFuori As Integer = 0
				Dim PareggiTorneiFuori As Integer = 0
				Dim SconfitteTorneiFuori As Integer = 0
				Dim VittorieTorneiCampoEsterno As Integer = 0
				Dim PareggiTorneiCampoEsterno As Integer = 0
				Dim SconfitteTorneiCampoEsterno As Integer = 0

				Dim SquadreIncontrate As New List(Of String)
				Dim MarcatoriGenerali As New List(Of String)
				Dim Presenze As New List(Of String)
				Dim MarcatoriTutte As New List(Of String)
				Dim PresenzeTutte As New List(Of String)

				Dim maxGoalInUnaPartita As Integer = 0
				Dim PartitaConPiuGoal As Integer = -1
				Dim minGoalInUnaPartita As Integer = 999
				Dim PartitaConMenoGoal As Integer = -1

				Dim TipologiaPartitePerAnno As String = ""

				Dim sPartitaConPiuGoal As String = ""
				Dim sPartitaConMenoGoal As String = ""
				Dim TempoTotaleDiGioco As String = ""

				Dim GoalSubitiPerMinuto As String = ""
				Dim GoalFattiPerMinuto As String = ""

				Dim ListaPartiteDiCampionato As String = ""
				Dim ListaPartiteAmichevoli As String = ""
				Dim ListaPartiteTornei As String = ""

				Dim ListaMarkers As String = ""
				Dim GraficoPuntiCampionato As String = ""
				Dim PuntiTotali As Integer = 0
				Dim PuntiTotaliMedia As Integer = 0
				Dim GiornataPerGrafico As Integer = 0

				Dim LatCasa As String = ""
				Dim LonCasa As String = ""

				Dim lmLat As New List(Of String)
				Dim lmLon As New List(Of String)
				Dim lmInd As New List(Of String)
				Dim lmDescr As New List(Of String)

				Dim descAnno As String = ""
				Dim nomeSquadra As String = ""

				Dim PathBaseImmagini As String = "http://loppa.duckdns.org:90/MultiMedia" ' "http://looigi.no-ip.biz:90/CVCalcio/App_Themes/Standard/Images"
				Dim PathBaseImmScon As String = "http://loppa.duckdns.org:90/MultiMedia/Sconosciuto.png" ' "http://looigi.no-ip.biz:90/CVCalcio/App_Themes/Standard/Images/Sconosciuto.png"
				Dim PathImmaginiPartite As String = "C:\GestioneCampionato\CalcioImages\" & Squadra & "\Partite\" ' "C:\inetpub\wwwroot\CVCalcio\App_Themes\Standard\Images\" & Squadra & "\Partite\"
				Dim pathImmagini As String = "http://loppa.duckdns.org:90/MultiMedia/"

				Sql = "SELECT * From Anni Where idAnno=" & idAnno
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof Then
							descAnno = Rec("Descrizione").Value
							nomeSquadra = Rec("NomeSquadra").Value
						End If

						' Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT TipologiePartite.Descrizione, Count(*) As Quante " &
					"FROM Partite Left Join TipologiePartite On Partite.idTipologia = TipologiePartite.idTipologia " &
					"Where Partite.idAnno = " & idAnno & " And Partite.idCategoria = " & idCategoria & "  " &
					"Group By TipologiePartite.Descrizione"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							TipologiaPartitePerAnno &= Rec("Descrizione").Value & " " & Rec("Quante").Value & "§"

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT Partite.idPartita, Partite.Casa " &
					"FROM Partite " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria & " " &
					"And Partite.idTipologia=1"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							PartiteCampionato.Add(Rec("idPartita").Value)
							PartiteCampionatoDove.Add(Rec("Casa").Value)
							PartiteCampionatoIN += Rec("idPartita").Value.ToString & ","

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT Partite.idPartita, Partite.Casa " &
					"FROM Partite " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria & " " &
					"And Partite.idTipologia=2"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							PartiteAmichevoli.Add(Rec("idPartita").Value)
							PartiteAmichevoliDove.Add(Rec("Casa").Value)
							PartiteAmichevoliIN += Rec("idPartita").Value.ToString & ","

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				' Tornei
				Sql = "SELECT Partite.idPartita, Partite.Casa " &
					"FROM Partite " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria & " " &
					"And Partite.idTipologia=3"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							PartiteTornei.Add(Rec("idPartita").Value)
							PartiteTorneiDove.Add(Rec("Casa").Value)
							PartiteTorneiIN += Rec("idPartita").Value.ToString & ","

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				If PartiteCampionatoIN.Length > 0 Then
					PartiteCampionatoIN = Mid(PartiteCampionatoIN, 1, PartiteCampionatoIN.Length - 1)
					Sql = "SELECT 'GoalCampionatoCasa' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' " &
						"Union All " &
						"SELECT 'RigoriCampionatoCasa' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' And Termine=1 " &
						"Union All " &
						"SELECT 'GoalCampionatoFuori' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' " &
						"Union All " &
						"SELECT 'RigoriCampionatoFuori' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' And Termine=1 " &
						"Union All " &
						"SELECT 'GoalCampionatoCampoEsterno' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E' " &
						"Union All " &
						"SELECT 'RigoriCampionatoCampoEsterno' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E' And Termine=1 "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "GoalCampionatoCasa"
										GoalCampionatoCasa += Rec(1).Value
									Case "GoalCampionatoFuori"
										GoalCampionatoFuori += Rec(1).Value
									Case "GoalCampionatoCasa"
										GoalCampionatoCampoEsterno += Rec(1).Value

									Case "RigoriCampionatoCasa"
										GoalCampionatoCasa += Rec(1).Value
									Case "RigoriCampionatoFuori"
										GoalCampionatoFuori += Rec(1).Value
									Case "RigoriCampionatoCasa"
										GoalCampionatoCampoEsterno += Rec(1).Value
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteTorneiIN.Length > 0 Then
					PartiteTorneiIN = Mid(PartiteTorneiIN, 1, PartiteTorneiIN.Length - 1)
					Sql = "SELECT 'GoalTorneiCasa' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' " &
						"Union All " &
						"SELECT 'RigoriTorneiCasa' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' And Termine=1 " &
						"Union All " &
						"SELECT 'GoalTorneiFuori' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' " &
						"Union All " &
						"SELECT 'RigoriTorneiFuori' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' And Termine=1 " &
						"Union All " &
						"SELECT 'GoalTorneiCampoEsterno' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E' " &
						"Union All " &
						"SELECT 'RigoriTorneiCampoEsterno' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E' And Termine=1 "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "GoalTorneiCasa"
										GoalTorneiCasa += Rec(1).Value
									Case "GoalTorneiFuori"
										GoalTorneiFuori += Rec(1).Value
									Case "GoalTorneiCampoEsterno"
										GoalTorneiCampoEsterno += Rec(1).Value
									Case "RigoriTorneiCasa"
										GoalTorneiCasa += Rec(1).Value
									Case "RigoriTorneiFuori"
										GoalTorneiFuori += Rec(1).Value
									Case "RigoriTorneiCampoEsterno"
										GoalTorneiCampoEsterno += Rec(1).Value
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteAmichevoliIN.Length > 0 Then
					PartiteAmichevoliIN = Mid(PartiteAmichevoliIN, 1, PartiteAmichevoliIN.Length - 1)
					Sql = "SELECT 'GoalAmichevoliCasa' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' " &
						"Union All " &
						"SELECT 'RigoriAmichevoliCasa' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' And Termine=1 " &
						"Union All " &
						"SELECT 'GoalAmichevoliFuori' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' " &
						"Union All " &
						"SELECT 'RigoriAmichevoliFuori' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' And Termine=1 " &
						"Union All " &
						"SELECT 'GoalAmichevoliCampoEsterno' As Cosa, Count(*) As GoalTotali " &
						"From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E' " &
						"Union All " &
						"SELECT 'RigoriAmichevoliCampoEsterno' As Cosa, Count(*) As GoalTotali " &
						"From RigoriPropri Left Join Partite On RigoriPropri.idPartita=Partite.idPartita " &
						"Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E' And Termine=1 "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "GoalAmichevoliCasa"
										GoalAmichevoliCasa += Rec(1).Value
									Case "GoalAmichevoliFuori"
										GoalAmichevoliFuori += Rec(1).Value
									Case "GoalAmichevoliCampoEsterno"
										GoalAmichevoliCampoEsterno += Rec(1).Value

									Case "RigoriAmichevoliCasa"
										GoalAmichevoliCasa += Rec(1).Value
									Case "RigoriAmichevoliFuori"
										GoalAmichevoliFuori += Rec(1).Value
									Case "RigoriAmichevoliCampoEsterno"
										GoalAmichevoliCampoEsterno += Rec(1).Value
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteCampionatoIN.Length > 0 Then
					Sql = "SELECT 'MarcatoriCasaCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
						"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
						"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'RigoriCasaCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
						"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
						"Where RigoriPropri.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' And Termine=1 And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'MarcatoriFuoriCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
						"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
						"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'RigoriFuoriCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
						"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
						"Where RigoriPropri.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' And Termine=1 And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'MarcatoriCampoEsternoCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
						"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
						"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E' And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'RigoriCampoEsternoCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
						"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
						"Where RigoriPropri.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E' And Termine=1 And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "MarcatoriCasaCampionato"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriCampionatoCasa.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											' NomiMarcatoriCampionatoCasa.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriCampionatoCasa = AggiungeRigoriEGoal(NomiMarcatoriCampionatoCasa, Rec)
										End If
										MarcatoriCampionatoCasa += Rec(3).Value
									Case "MarcatoriFuoriCampionato"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriCampionatoFuori.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											' NomiMarcatoriCampionatoFuori.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriCampionatoFuori = AggiungeRigoriEGoal(NomiMarcatoriCampionatoFuori, Rec)
										End If
										MarcatoriCampionatoFuori += Rec(3).Value
									Case "MarcatoriCampoEsternoCampionato"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriCampionatoCampoEsterno.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											'NomiMarcatoriCampionatoCampoEsterno.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriCampionatoCampoEsterno = AggiungeRigoriEGoal(NomiMarcatoriCampionatoCampoEsterno, Rec)
										End If
										MarcatoriCampionatoCampoEsterno += Rec(3).Value

									Case "RigoriCasaCampionato"
										NomiMarcatoriCampionatoCasa = AggiungeRigoriEGoal(NomiMarcatoriCampionatoCasa, Rec)

										MarcatoriCampionatoCasa += Rec(3).Value
									Case "RigoriFuoriCampionato"
										NomiMarcatoriCampionatoFuori = AggiungeRigoriEGoal(NomiMarcatoriCampionatoFuori, Rec)

										MarcatoriCampionatoFuori += Rec(3).Value
									Case "RigoriCampoEsternoCampionato"
										NomiMarcatoriCampionatoCampoEsterno = AggiungeRigoriEGoal(NomiMarcatoriCampionatoCampoEsterno, Rec)

										MarcatoriCampionatoCampoEsterno += Rec(3).Value
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteAmichevoliIN.Length > 0 Then
					Sql = "SELECT 'MarcatoriCasaAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
					"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
					"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
					"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' " &
					"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
					"Union All " &
					"Select 'RigoriCasaAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
					"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
					"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
					"Where RigoriPropri.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' And Termine=1 " &
					"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
					"Union All " &
					"Select 'MarcatoriFuoriAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
					"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
					"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
					"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' " &
					"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
					"Union All " &
					"Select 'RigoriFuoriAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
					"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
					"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
					"Where RigoriPropri.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' And Termine=1 " &
					"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
					"Union All " &
					"Select 'MarcatoriCampoEsternoAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
					"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
					"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
					"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E' " &
					"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
					"Union All " &
					"Select 'RigoriCampoEsternoAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
					"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
					"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
					"Where RigoriPropri.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E' And Termine=1 " &
					"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "MarcatoriCasaAmichevoli"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriAmichevoliCasa.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											' NomiMarcatoriAmichevoliCasa.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriAmichevoliCasa = AggiungeRigoriEGoal(NomiMarcatoriAmichevoliCasa, Rec)
										End If
										MarcatoriAmichevoliCasa += Rec(3).Value
									Case "MarcatoriFuoriAmichevoli"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriAmichevoliFuori.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											' NomiMarcatoriAmichevoliFuori.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriAmichevoliFuori = AggiungeRigoriEGoal(NomiMarcatoriAmichevoliFuori, Rec)
										End If
										MarcatoriAmichevoliFuori += Rec(3).Value
									Case "MarcatoriCampoEsternoAmichevoli"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriAmichevoliCampoEsterno.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											' NomiMarcatoriAmichevoliCampoEsterno.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriAmichevoliCampoEsterno = AggiungeRigoriEGoal(NomiMarcatoriAmichevoliCampoEsterno, Rec)
										End If
										MarcatoriAmichevoliCampoEsterno += Rec(3).Value

									Case "RigoriCasaAmichevoli"
										NomiMarcatoriAmichevoliCasa = AggiungeRigoriEGoal(NomiMarcatoriAmichevoliCasa, Rec)

										MarcatoriAmichevoliCasa += Rec(3).Value
									Case "RigoriFuoriAmichevoli"
										NomiMarcatoriAmichevoliFuori = AggiungeRigoriEGoal(NomiMarcatoriAmichevoliFuori, Rec)

										MarcatoriAmichevoliFuori += Rec(3).Value
									Case "RigoriCampoEsternoAmichevoli"
										NomiMarcatoriAmichevoliCampoEsterno = AggiungeRigoriEGoal(NomiMarcatoriAmichevoliCampoEsterno, Rec)

										MarcatoriAmichevoliCampoEsterno += Rec(3).Value
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteTorneiIN.Length > 0 Then
					Sql = "SELECT 'MarcatoriCasaTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
						"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
						"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'RigoriCasaTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
						"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
						"Where RigoriPropri.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' And Termine=1 And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'MarcatoriFuoriTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
						"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
						"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'RigoriFuoriTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
						"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
						"Where RigoriPropri.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' And Termine=1 And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'MarcatoriCampoEsternoTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
						"Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
						"Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E' " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore " &
						"Union All " &
						"Select 'RigoriCampoEsternoTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
						"FROM(RigoriPropri Left Join Partite On Partite.idPartita = RigoriPropri.idPartita) " &
						"Left Join Giocatori On RigoriPropri.idGiocatore = Giocatori.idGiocatore " &
						"Where RigoriPropri.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E' And Termine=1 And Giocatori.idAnno=Partite.idAnno " &
						"Group By Giocatori.Cognome, Giocatori.Nome, Giocatori.idGiocatore "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "MarcatoriCasaTornei"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriTorneiCasa.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											' NomiMarcatoriTorneiCasa.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriTorneiCasa = AggiungeRigoriEGoal(NomiMarcatoriTorneiCasa, Rec)
										End If

										MarcatoriTorneiCasa += Rec(3).Value
									Case "MarcatoriFuoriTornei"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriTorneiFuori.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											' NomiMarcatoriTorneiFuori.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriTorneiFuori = AggiungeRigoriEGoal(NomiMarcatoriTorneiFuori, Rec)
										End If

										MarcatoriTorneiFuori += Rec(3).Value
									Case "MarcatoriCampoEsternoTornei"
										If "" & Rec(1).Value = "" Or "" & Rec(2).Value = "" Then
											NomiMarcatoriTorneiCampoEsterno.Add(Rec(4).Value & "-Autorete-" & Rec(3).Value)
										Else
											'NomiMarcatoriTorneiCampoEsterno.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
											NomiMarcatoriTorneiCampoEsterno = AggiungeRigoriEGoal(NomiMarcatoriTorneiCampoEsterno, Rec)
										End If

										MarcatoriTorneiCampoEsterno += Rec(3).Value

									Case "RigoriCasaTornei"
										NomiMarcatoriTorneiCasa = AggiungeRigoriEGoal(NomiMarcatoriTorneiCasa, Rec)

										MarcatoriTorneiCasa += Rec(3).Value
									Case "RigoriFuoriTornei"
										NomiMarcatoriTorneiFuori = AggiungeRigoriEGoal(NomiMarcatoriTorneiFuori, Rec)

										MarcatoriTorneiFuori += Rec(3).Value
									Case "RigoriCampoEsternoTornei"
										NomiMarcatoriTorneiCampoEsterno = AggiungeRigoriEGoal(NomiMarcatoriTorneiCampoEsterno, Rec)

										MarcatoriTorneiCampoEsterno += Rec(3).Value
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteCampionatoIN.Length > 0 Then
					Sql = "SELECT 'AvversariCasa' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
						"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
						"Where RisultatiAggiuntivi.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' " &
						"Union All " &
						"SELECT 'AvversariRigoriCasa' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
						"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
						"Where RigoriAvversari.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' And Giocatori.idAnno=Partite.idAnno " &
						"Union All " &
						"SELECT 'AvversariFuori' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
						"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
						"Where RisultatiAggiuntivi.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' " &
						"Union All " &
						"SELECT 'AvversariRigoriFuori' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
						"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
						"Where RigoriAvversari.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' And Giocatori.idAnno=Partite.idAnno" &
						"Union All " &
						"SELECT 'AvversariCampoEsterno' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
						"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
						"Where RisultatiAggiuntivi.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E' " &
						"Union All " &
						"SELECT 'AvversariRigoriCampoEsterno' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
						"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
						"Where RigoriAvversari.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E' And Giocatori.idAnno=Partite.idAnno "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "AvversariCasa"
										GoalAvvCampionatoCasa1Tempo = Val("" & Rec(1).Value)
										GoalAvvCampionatoCasa2Tempo = Val("" & Rec(2).Value)
										GoalAvvCampionatoCasa3Tempo += Val("" & Rec(3).Value)
									Case "AvversariFuori"
										GoalAvvCampionatoFuori1Tempo = Val("" & Rec(1).Value)
										GoalAvvCampionatoFuori2Tempo = Val("" & Rec(2).Value)
										GoalAvvCampionatoFuori3Tempo += Val("" & Rec(3).Value)
									Case "AvversariCampoEsterno"
										GoalAvvCampionatoCampoEsterno1Tempo = Val("" & Rec(1).Value)
										GoalAvvCampionatoCampoEsterno2Tempo = Val("" & Rec(2).Value)
										GoalAvvCampionatoCampoEsterno3Tempo += Val("" & Rec(3).Value)

									Case "AvversariRigoriCasa"
										GoalAvvCampionatoCasa3Tempo += Val("" & Rec(1).Value)
									Case "AvversariRigoriFuori"
										GoalAvvCampionatoFuori3Tempo += Val("" & Rec(1).Value)
									Case "AvversariRigoriCampoEsterno"
										GoalAvvCampionatoCampoEsterno3Tempo += Val("" & Rec(1).Value)
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteAmichevoliIN.Length > 0 Then
					Sql = "SELECT 'AvversariCasa' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
					"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
					"Where RisultatiAggiuntivi.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' " &
					"Union All " &
					"SELECT 'AvversariRigoriCasa' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
					"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
					"Where RigoriAvversari.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' And Giocatori.idAnno=Partite.idAnno " &
					"Union All " &
					"SELECT 'AvversariFuori' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
					"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
					"Where RisultatiAggiuntivi.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' " &
					"Union All " &
					"SELECT 'AvversariRigoriFuori' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
					"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
					"Where RigoriAvversari.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' And Giocatori.idAnno=Partite.idAnno " &
					"Union All " &
					"SELECT 'AvversariCampoEsterno' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
					"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
					"Where RisultatiAggiuntivi.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E' " &
					"Union All " &
					"SELECT 'AvversariRigoriCampoEsterno' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
					"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
					"Where RigoriAvversari.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E' And Giocatori.idAnno=Partite.idAnno "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "AvversariCasa"
										GoalAvvAmichevoliCasa1Tempo = Val("" & Rec(1).Value)
										GoalAvvAmichevoliCasa2Tempo = Val("" & Rec(2).Value)
										GoalAvvAmichevoliCasa3Tempo += Val("" & Rec(3).Value)
									Case "AvversariFuori"
										GoalAvvAmichevoliFuori1Tempo = Val("" & Rec(1).Value)
										GoalAvvAmichevoliFuori2Tempo = Val("" & Rec(2).Value)
										GoalAvvAmichevoliFuori3Tempo += Val("" & Rec(3).Value)
									Case "AvversariCampoEsterno"
										GoalAvvAmichevoliCampoEsterno1Tempo = Val("" & Rec(1).Value)
										GoalAvvAmichevoliCampoEsterno2Tempo = Val("" & Rec(2).Value)
										GoalAvvAmichevoliCampoEsterno3Tempo += Val("" & Rec(3).Value)

									Case "AvversariRigoriCasa"
										GoalAvvAmichevoliCasa3Tempo += Val("" & Rec(1).Value)
									Case "AvversariRigoriFuori"
										GoalAvvAmichevoliFuori3Tempo += Val("" & Rec(1).Value)
									Case "AvversariRigoriCampoEsterno"
										GoalAvvAmichevoliCampoEsterno3Tempo += Val("" & Rec(1).Value)
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				If PartiteTorneiIN.Length > 0 Then
					Sql = "SELECT 'AvversariCasa' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
					"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
					"Where RisultatiAggiuntivi.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' " &
					"Union All " &
					"SELECT 'AvversariRigoriCasa' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
					"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
					"Where RigoriAvversari.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' And Giocatori.idAnno=Partite.idAnno " &
					"Union All " &
					"SELECT 'AvversariFuori' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
					"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
					"Where RisultatiAggiuntivi.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' " &
					"Union All " &
					"SELECT 'AvversariRigoriFuori' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
					"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
					"Where RigoriAvversari.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' And Giocatori.idAnno=Partite.idAnno " &
					"Union All " &
					"SELECT 'AvversariCampoEsterno' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
					"From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
					"Where RisultatiAggiuntivi.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E' " &
					"Union All " &
					"SELECT 'AvversariRigoriCampoEsterno' As Cosa, Sum(Segnati) As PrimoTempo, 0 As SecondoTempo, 0 As TerzoTempo " &
					"From RigoriAvversari Left Join Partite On RigoriAvversari.idPartita = Partite.idPartita " &
					"Where RigoriAvversari.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E' And Giocatori.idAnno=Partite.idAnno "
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof
								Select Case Rec("Cosa").Value
									Case "AvversariCasa"
										GoalAvvTorneiCasa1Tempo = Val("" & Rec(1).Value)
										GoalAvvTorneiCasa2Tempo = Val("" & Rec(2).Value)
										GoalAvvTorneiCasa3Tempo += Val("" & Rec(3).Value)
									Case "AvversariFuori"
										GoalAvvTorneiFuori1Tempo = Val("" & Rec(1).Value)
										GoalAvvTorneiFuori2Tempo = Val("" & Rec(2).Value)
										GoalAvvTorneiFuori3Tempo += Val("" & Rec(3).Value)
									Case "AvversariCampoEsterno"
										GoalAvvTorneiCampoEsterno1Tempo = Val("" & Rec(1).Value)
										GoalAvvTorneiCampoEsterno2Tempo = Val("" & Rec(2).Value)
										GoalAvvTorneiCampoEsterno3Tempo += Val("" & Rec(3).Value)

									Case "AvversariRigoriCasa"
										GoalAvvTorneiCasa3Tempo += Val("" & Rec(1).Value)
									Case "AvversariRigoriFuori"
										GoalAvvTorneiFuori3Tempo += Val("" & Rec(1).Value)
									Case "AvversariRigoriCampoEsterno"
										GoalAvvTorneiCampoEsterno3Tempo += Val("" & Rec(1).Value)
								End Select

								Rec.MoveNext
							Loop
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				Sql = "SELECT Partite.idAvversario, Descrizione, Count(*) As Quante From " &
					"Partite Left Join SquadreAvversarie On Partite.idAvversario = SquadreAvversarie.idAvversario " &
					"Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " " &
					"Group By Partite.idAvversario, Descrizione " &
					"Order By 3 Desc"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							SquadreIncontrate.Add(Rec("idAvversario").Value & ";" & Rec("Descrizione").Value & ";" & Rec("Quante").Value)

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "Select Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Sum(Quanti2) As Quanti From (" &
					"SELECT Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) As Quanti2 " &
					"FROM (RisultatiAggiuntiviMarcatori INNER JOIN Partite On RisultatiAggiuntiviMarcatori.idPartita = Partite.idPartita) INNER JOIN Giocatori On (RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore) And (Partite.idAnno = Giocatori.idAnno) " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria & " " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc " &
					"Union All " &
					"Select Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) As Quanti2 " &
					"FROM (RigoriPropri INNER JOIN Partite On RigoriPropri.idPartita = Partite.idPartita) INNER JOIN Giocatori On (RigoriPropri.idGiocatore = Giocatori.idGiocatore) And (Partite.idAnno = Giocatori.idAnno) " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria & " And RigoriPropri.Termine=1 And RigoriPropri.idAnno=" & idAnno & " And Giocatori.idAnno=" & idAnno & " " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc " &
					") A " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							If "" & Rec("Cognome").Value = "" Or "" & Rec("Nome").Value = "" Then
								MarcatoriGenerali.Add(Rec("idGiocatore").Value & ";Autorete;" & Rec("Quanti").Value)
							Else
								MarcatoriGenerali.Add(Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & " " & Rec("Nome").Value & ";" & Rec("Quanti").Value)
							End If

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "Select Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Sum(Quanti2) As Quanti From (" &
					"SELECT Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) As Quanti2 " &
					"FROM (RisultatiAggiuntiviMarcatori INNER JOIN Partite On RisultatiAggiuntiviMarcatori.idPartita = Partite.idPartita) INNER JOIN Giocatori On (RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore) And (Partite.idAnno = Giocatori.idAnno) " &
					"WHERE Partite.idAnno=" & idAnno & " " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc " &
					"Union All " &
					"Select Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) As Quanti2 " &
					"FROM (RigoriPropri INNER JOIN Partite On RigoriPropri.idPartita = Partite.idPartita) INNER JOIN Giocatori On (RigoriPropri.idGiocatore = Giocatori.idGiocatore) And (Partite.idAnno = Giocatori.idAnno) " &
					"WHERE Partite.idAnno=" & idAnno & " And RigoriPropri.Termine=1 And RigoriPropri.idAnno=" & idAnno & " And Giocatori.idAnno=" & idAnno & " " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc " &
					") A " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							If "" & Rec("Cognome").Value = "" Or "" & Rec("Nome").Value = "" Then
								MarcatoriTutte.Add(Rec("idGiocatore").Value & ";Autorete;" & Rec("Quanti").Value)
							Else
								MarcatoriTutte.Add(Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & " " & Rec("Nome").Value & ";" & Rec("Quanti").Value)
							End If

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) As Presenze " &
					"FROM (Partite LEFT JOIN Convocati ON Partite.idPartita = Convocati.idPartita) LEFT JOIN Giocatori ON Convocati.idGiocatore = Giocatori.idGiocatore " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria & " And Giocatori.idAnno=" & idAnno & " " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							Presenze.Add(Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & " " & Rec("Nome").Value & ";" & Rec("Presenze").Value)

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Count(*) As Presenze " &
					"FROM (Partite INNER JOIN Convocati ON Partite.idPartita = Convocati.idPartita) INNER JOIN Giocatori ON (Partite.idAnno = Giocatori.idAnno) AND (Convocati.idGiocatore = Giocatori.idGiocatore) " &
					"WHERE Partite.idAnno=" & idAnno & " " &
					"Group By Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome " &
					"Order By 4 Desc"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							PresenzeTutte.Add(Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & " " & Rec("Nome").Value & ";" & Rec("Presenze").Value)

							Rec.MoveNext
						Loop
						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT Minuto, Count(*) As Quanti " &
					"FROM RisultatiAggiuntiviMarcatori INNER JOIN Partite ON RisultatiAggiuntiviMarcatori.idPartita = Partite.idPartita " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria & " " &
					"Group By Minuto " &
					"Order By 2 Desc"
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Dim Stringozza As String = ""

						Stringozza &= "<table cellspacing=""0"" style=""width: 100%;"">"
						Do Until Rec.Eof
							Stringozza &= "<tr " & RitornaColoreSfondo() & ">"
							Stringozza &= "<td style=""width: 50%; text-align: center;"">"
							Stringozza &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("Minuto").Value & "°</span>"
							Stringozza &= "</td>"
							Stringozza &= "<td style=""width: 50%; text-align: center;"">"
							Stringozza &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("Quanti").Value & "</span>"
							Stringozza &= "</td>"
							Stringozza &= "</tr>"

							Rec.MoveNext
						Loop
						Stringozza &= "</table>"

						GoalFattiPerMinuto = Stringozza

						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT TempiGoalAvversari.TempiPrimoTempo, TempiGoalAvversari.TempiSecondoTempo, TempiGoalAvversari.TempiTerzoTempo " &
					"FROM TempiGoalAvversari LEFT JOIN Partite ON TempiGoalAvversari.idPartita = Partite.idPartita " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Dim Stringozza As String = ""
						Dim Minuti(60) As Integer
						Dim sMinuti(60) As String

						Do Until Rec.Eof
							If "" & Rec("TempiPrimoTempo").Value <> "" Then
								Dim t() As String = Rec("TempiPrimoTempo").Value.split("#")

								For Each Minuto As String In t
									If Minuto <> "" Then
										Minuti(Val(Minuto)) += 1
										sMinuti(Val(Minuto)) = Minuto
									End If
								Next
							End If
							If "" & Rec("TempiSecondoTempo").Value <> "" Then
								Dim t() As String = Rec("TempiSecondoTempo").Value.split("#")

								For Each Minuto As String In t
									If Minuto <> "" Then
										Minuti(Val(Minuto)) += 1
										sMinuti(Val(Minuto)) = Minuto
									End If
								Next
							End If
							If "" & Rec("TempiTerzoTempo").Value <> "" Then
								Dim t() As String = Rec("TempiTerzoTempo").Value.split("#")

								For Each Minuto As String In t
									If Minuto <> "" Then
										Minuti(Val(Minuto)) += 1
										sMinuti(Val(Minuto)) = Minuto
									End If
								Next
							End If

							Rec.MoveNext
						Loop

						For i As Integer = 0 To 60
							For k As Integer = 0 To 60
								If Minuti(i) > Minuti(k) Then
									Dim Appoggio As Integer = Minuti(i)
									Minuti(i) = Minuti(k)
									Minuti(k) = Appoggio

									Dim sAppoggio As String = sMinuti(i)
									sMinuti(i) = sMinuti(k)
									sMinuti(k) = sAppoggio
								End If
							Next
						Next

						Stringozza &= "<table cellspacing=""0"" style=""width: 100%;"">"
						For i As Integer = 0 To 60
							If Minuti(i) > 0 Then
								Stringozza &= "<tr " & RitornaColoreSfondo() & ">"
								Stringozza &= "<td style=""width: 50%; text-align: center;"">"
								Stringozza &= "<span class=""testo nero"" style=""font-size: 16px;"">" & sMinuti(i) & "°</span>"
								Stringozza &= "</td>"
								Stringozza &= "<td style=""width: 50%; text-align: center;"">"
								Stringozza &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Minuti(i) & "</span>"
								Stringozza &= "</td>"
								Stringozza &= "</tr>"
							End If
						Next
						Stringozza &= "</table>"

						GoalSubitiPerMinuto = Stringozza

						Rec.Close()
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				GraficoPuntiCampionato &= "['0', 0, 0], "
				For i As Integer = 1 To 3
					Sql = "SELECT Partite.idPartita, Partite.RisultatoATempi, Partite.DataOra, Allenatori.Cognome+ '<br />' +Allenatori.Nome As Allenatore,MeteoPartite.Tempo, MeteoPartite.Gradi, " &
						"Partite.idCategoria, SquadreAvversarie.idAvversario, SquadreAvversarie.Descrizione, " &
						"(Select Count(*) From RisultatiAggiuntiviMarcatori Where idPartita = Partite.idPartita) As Goal, " &
						"(Select (iif(GoalAvvPrimoTempo>-1,GoalAvvPrimoTempo,0)+iif(GoalAvvSecondoTempo>-1,GoalAvvSecondoTempo,0)+iif(GoalAvvTerzoTempo>-1,GoalAvvTerzoTempo,0)) " &
						"From RisultatiAggiuntivi Where idPartita = Partite.idPartita) As GoalAvv, " &
						"RisGiochetti, Arbitri.Cognome+ '<br />'+Arbitri.Nome As Arbitro, Anni.NomeSquadra, Partite.Casa, " &
						"RisultatiAggiuntivi.GoalAvvSecondoTempo, RisultatiAggiuntivi.GoalAvvPrimoTempo, RisultatiAggiuntivi.GoalAvvTerzoTempo, " &
						"(SELECT Count(*) FROM RisultatiAggiuntiviMarcatori Where idPartita = Partite.idPartita And idTempo=1) As GoalPrimoTempo, " &
						"(SELECT Count(*) FROM RisultatiAggiuntiviMarcatori Where idPartita = Partite.idPartita And idTempo=2) As GoalSecondoTempo, " &
						"(SELECT Count(*) FROM RisultatiAggiuntiviMarcatori Where idPartita = Partite.idPartita And idTempo=3) As GoalTerzoTempo, " &
						"CampiAvversari.Descrizione As Campo, CampiAvversari.Indirizzo As IndirizzoCampo, Partite.idPartita, Anni.Indirizzo As IndirizzoCasa, " &
						"Anni.Lat As LatCasa, Anni.Lon As LonCasa " &
						"FROM (((((((Partite LEFT JOIN Allenatori ON Partite.idAllenatore = Allenatori.idAllenatore And Allenatori.idAnno=Partite.idAnno) " &
						"LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) " &
						"LEFT JOIN MeteoPartite ON Partite.idPartita = MeteoPartite.idPartita) " &
						"Left Join RisultatiAggiuntivi On Partite.idPartita=RisultatiAggiuntivi.idPartita) " &
						"LEFT JOIN ArbitriPartite On Partite.idPartita = ArbitriPartite.idPartita) " &
						"LEFT JOIN Arbitri On ArbitriPartite.idArbitro=Arbitri.idArbitro) " &
						"LEFT JOIN Anni On Partite.idAnno = Anni.idAnno) " &
						"LEFT JOIN CampiAvversari On SquadreAvversarie.idCampo = CampiAvversari.idCampo " &
						"WHERE Partite.idAnno=" & idAnno & " And Partite.idCategoria=" & idCategoria & " And Partite.idTipologia=" & i & " And Arbitri.idAnno=Partite.idAnno " &
						"Order By DataOra"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Not Rec.Eof Then
								Dim Stringozza As String = ""

								LatCasa = Rec("LatCasa").Value
								LonCasa = Rec("LonCasa").Value

								Do Until Rec.Eof
									gf.ScansionaDirectorySingola(PathImmaginiPartite & Rec("idPartita").Value.ToString.Trim)
									Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
									Dim Filetti() As String = gf.RitornaFilesRilevati

									Stringozza &= "<table cellspacing=""0"" style=""width: 100%;""> "
									Stringozza &= "<tr>"
									Stringozza &= "<td colspan=""11"" style=""background-color: #555; height: 5px;""></td>"
									Stringozza &= "</tr>"
									Stringozza &= "<tr>"
									Stringozza &= "<td style=""text-align: center; background-color: #ccc; margin: 3px; width: 55px;""> "
									Stringozza &= "<img src=""" & pathImmagini & "icona_CERCA.png"" "
									Stringozza &= "style=""width: 50px; height: 50px; cursor: pointer;"""
									Stringozza &= "onclick=""VisualizzaPartita('" & idAnno & "_" & Rec("idPartita").Value & "');"" "
									Stringozza &= "title=""Visualizza dettaglio partita " & Rec("idPartita").Value.ToString.Trim & """ /> "
									If qFiles > 0 Then
										Dim sFiles As String = ""
										For ii As Integer = 1 To qFiles
											Dim nomeF As String = gf.TornaNomeFileDaPath(Filetti(ii))
											nomeF = pathImmagini & Squadra & "/Partite/" & Rec("idPartita").Value.ToString.Trim & "/" & nomeF
											sFiles &= nomeF & "§"
										Next
										If sFiles <> "" Then
											sFiles = Mid(sFiles, 1, sFiles.Length - 1)
										End If

										Stringozza &= "<hr />"
										Stringozza &= "<img src=""" & pathImmagini & "Icone/visualizzato_tondo.png"" "
										Stringozza &= "style=""width: 50px; height: 50px; cursor: pointer;"""
										Stringozza &= "onclick=""VisualizzaImmagini('" & sFiles & "');"" "
										Stringozza &= "title=""Visualizza multimedia della partita " & Rec("idPartita").Value.ToString.Trim & """ /> "
									End If
									Stringozza &= "</td>"
									Stringozza &= "<td>"

									Stringozza &= "<table cellspacing=""0"" style=""width: 100%;"">"
									' Stringozza &= "<tr>"
									' Stringozza &= "<td colspan=""11"" style=""background-color: #555; height: 5px;""></td>"
									' Stringozza &= "</tr>"

									Stringozza &= "<tr " & RitornaColoreSfondo() & ">"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">" & Rec("DataOra").Value.ToString.Replace(" ", "<br />") & "</span>"
									Stringozza &= "</td>"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">" & Rec("Allenatore").Value & "</span>"
									Stringozza &= "</td>"

									Dim Tempo As String = ""

									If "" & Rec("Tempo").Value <> "" Then
										Tempo = Rec("Tempo").Value
									End If
									If "" & Rec("Gradi").Value <> "" Then
										Tempo &= "<br />" & Rec("Gradi").Value & " Gradi"
									End If
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">" & Tempo & "</span>"
									Stringozza &= "</td>"

									Dim Casa As String
									Dim Fuori As String
									Dim Imm1 As String = PathBaseImmagini & "/" & Squadra & "/Categorie/" & idAnno & "_" & idCategoria & ".Jpg"
									Dim Imm2 As String = PathBaseImmagini & "/Avversari/" & Rec("idAvversario").Value & ".Jpg"
									Dim ImmCasa As String
									Dim ImmFuori As String
									Dim GoalCasa As String
									Dim GoalFuori As String

									Dim RigoriPropri As Integer = 0
									Dim RigoriAvversari As Integer = 0

									Try
										Sql = "Select Count(*) As Quanti From RigoriPropri Where idPartita=" & Rec("idPartita").Value & " And Termine=1"
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Not Rec2.Eof Then
												RigoriPropri = Rec2("Quanti").Value
											End If
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & " " & ex.Message
									End Try

									Try
										Sql = "Select Segnati As Quanti From RigoriAvversari Where idPartita=" & Rec("idPartita").Value
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Not Rec2.Eof Then
												RigoriAvversari = Rec2("Quanti").Value
											End If
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & " " & ex.Message
									End Try

									If Rec("Casa").Value = "S" Then
										Casa = Rec("NomeSquadra").Value
										Fuori = Rec("Descrizione").Value
										ImmCasa = "<td style =""text-align: center;""><img src=""" & Imm1 & """ style=""width: 55px; height: 55px; border: 1px solid #999;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
										ImmFuori = "<td style =""text-align: center;""><img src=""" & Imm2 & """ style=""width: 55px; height: 55px; border: 1px solid #999;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
										GoalCasa = Rec("Goal").Value
										GoalFuori = Rec("GoalAvv").Value

										GoalCasa += RigoriPropri
										GoalFuori += RigoriAvversari
									Else
										Fuori = Rec("NomeSquadra").Value
										Casa = Rec("Descrizione").Value
										ImmCasa = "<td style =""text-align: center;""><img src=""" & Imm2 & """ style=""width: 55px; height: 55px; border: 1px solid #999;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
										ImmFuori = "<td style =""text-align: center;""><img src=""" & Imm1 & """ style=""width: 55px; height: 55px; border: 1px solid #999;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
										GoalCasa = Rec("GoalAvv").Value
										GoalFuori = Rec("Goal").Value

										GoalCasa += RigoriAvversari
										GoalFuori += RigoriPropri
									End If
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo rosso"" style=""font-size: 15px;"">" & Casa.Replace(" ", "<br />") & "</span>"
									Stringozza &= "</td>"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= ImmCasa
									Stringozza &= "</td>"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">" & GoalCasa & "-" & GoalFuori & "</span>"
									Stringozza &= "</td>"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= ImmFuori
									Stringozza &= "</td>"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo rosso"" style=""font-size: 15px;"">" & Fuori.Replace(" ", "<br />") & "</span>"
									Stringozza &= "</td>"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">" & Rec("Arbitro").Value & "</span>"
									Stringozza &= "</td>"
									Stringozza &= "</tr>"

									Dim ATempi As String = Rec("RisultatoATempi").Value
									Dim GoalPrimoTempoCasa As Integer
									Dim GoalPrimoTempoFuori As Integer
									Dim GoalSecondoTempoCasa As Integer
									Dim GoalSecondoTempoFuori As Integer
									Dim GoalTerzoTempoCasa As Integer
									Dim GoalTerzoTempoFuori As Integer

									If Rec("Casa").Value = "S" Then
										GoalPrimoTempoCasa = Rec("GoalPrimoTempo").Value
										GoalSecondoTempoCasa = Rec("GoalSecondoTempo").Value
										GoalTerzoTempoCasa = Rec("GoalTerzoTempo").Value
										GoalPrimoTempoFuori = Rec("GoalAvvPrimoTempo").Value
										GoalSecondoTempoFuori = Rec("GoalAvvSecondoTempo").Value
										GoalTerzoTempoFuori = Rec("GoalAvvTerzoTempo").Value
									Else
										GoalPrimoTempoCasa = Rec("GoalAvvPrimoTempo").Value
										GoalSecondoTempoCasa = Rec("GoalAvvSecondoTempo").Value
										GoalTerzoTempoCasa = Rec("GoalAvvTerzoTempo").Value
										GoalPrimoTempoFuori = Rec("GoalPrimoTempo").Value
										GoalSecondoTempoFuori = Rec("GoalSecondoTempo").Value
										GoalTerzoTempoFuori = Rec("GoalTerzoTempo").Value
									End If

									Dim Punti1 As Integer = 0
									Dim Punti2 As Integer = 0

									If Rec("RisultatoATempi").Value = "S" Then
										If GoalPrimoTempoCasa > GoalPrimoTempoFuori Then
											Punti1 += 1
										Else
											If GoalPrimoTempoCasa < GoalPrimoTempoFuori Then
												Punti2 += 1
											Else
												Punti1 += 1
												Punti2 += 1
											End If
										End If
										If GoalSecondoTempoCasa > GoalSecondoTempoFuori Then
											Punti1 += 1
										Else
											If GoalSecondoTempoCasa < GoalSecondoTempoFuori Then
												Punti2 += 1
											Else
												Punti1 += 1
												Punti2 += 1
											End If
										End If
										If GoalTerzoTempoCasa > GoalTerzoTempoFuori Then
											Punti1 += 1
										Else
											If GoalTerzoTempoCasa < GoalTerzoTempoFuori Then
												Punti2 += 1
											Else
												Punti1 += 1
												Punti2 += 1
											End If
										End If

										Dim Giochetti As String = "" & Rec("RisGiochetti").Value

										If Giochetti <> "" Then
											Dim g() As String

											If Giochetti.Contains("-") Then
												g = Giochetti.Split("-")
											Else
												If Giochetti.Contains("/") Then
													g = Giochetti.Split("/")
												Else
													Stop
												End If
											End If

											If Val(g(0).Trim) > Val(g(1).Trim) Then
												'If Rec("Casa").Value = "S" Then
												Punti1 += 1
												'Else
												'    Punti2 += 1
												'End If
											Else
												If Val(g(0).Trim) < Val(g(1).Trim) Then
													'If Rec("Casa").Value = "S" Then
													Punti2 += 1
													'Else
													'    Punti1 += 1
													'End If
												Else
													Punti1 += 1
													Punti2 += 1
												End If
											End If
										End If
									Else
										Punti1 = GoalPrimoTempoCasa + GoalSecondoTempoCasa + GoalTerzoTempoCasa
										Punti2 = GoalPrimoTempoFuori + GoalSecondoTempoFuori + GoalTerzoTempoFuori
									End If

									Dim Esito As String = ""
									Dim Colore As String

									If Punti1 > Punti2 Then
										If Rec("Casa").Value = "S" Then
											Esito = "Vittoria"
											Colore = "#0a0"
											If i = 1 Then
												PuntiTotali += 3
											End If
										Else
											Esito = "Sconfitta"
											Colore = "#a00"
										End If
									Else
										If Punti1 < Punti2 Then
											If Rec("Casa").Value = "S" Then
												Esito = "Sconfitta"
												Colore = "#a00"
											Else
												If i = 1 Then
													PuntiTotali += 3
												End If
												Esito = "Vittoria"
												Colore = "#0a0"
											End If
										Else
											If i = 1 Then
												PuntiTotali += 1
											End If
											Esito = "Pareggio"
											Colore = "#000"
										End If
									End If
									If Rec("Casa").Value = "S" Then
										Esito &= " in casa"
									Else
										Esito &= " fuori casa"
									End If

									If i = 1 Then
										If Rec("Casa").Value = "S" Then
											PuntiTotaliMedia += 3
										Else
											PuntiTotaliMedia += 1
										End If
										GiornataPerGrafico += 1
										GraficoPuntiCampionato &= "['" & GiornataPerGrafico & "', " & PuntiTotali & ", " & PuntiTotaliMedia & "], "
									End If

									Stringozza &= "<tr " & RitornaColoreSfondo() & ">"
									Stringozza &= "<td style=""text-align: center;"">"
									Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">" & IIf(ATempi = "S", "Risultato<br /> a tempi", "Risultato<br />a goal") & "</span>"
									Stringozza &= "</td>"
									If Rec("RisultatoATempi").Value = "S" Then
										Stringozza &= "<td style=""text-align: center;"">"
										Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">1° Tempo: " & GoalPrimoTempoCasa & "-" & GoalPrimoTempoFuori & "</span>"
										Stringozza &= "</td>"
										Stringozza &= "<td style=""text-align: center;"">"
										Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">2° Tempo: " & GoalSecondoTempoCasa & "-" & GoalSecondoTempoFuori & "</span>"
										Stringozza &= "</td>"
										Stringozza &= "<td style=""text-align: center;"">"
										Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">3° Tempo: " & GoalTerzoTempoCasa & "-" & GoalTerzoTempoFuori & "</span>"
										Stringozza &= "</td>"
										Stringozza &= "<td style=""text-align: center;"" colspan=""2"">"
										Stringozza &= "<span class=""testo nero"" style=""font-size: 15px;"">Giochetti: " & Rec("RisGiochetti").Value & "</span>"
										Stringozza &= "</td>"
									Else
										Stringozza &= "<td colspan=""5"">"
										Stringozza &= "</td>"
									End If
									Stringozza &= "<td style=""text-align: center;"" colspan=""2"">"
									Stringozza &= "<span class=""testo blu"" style=""font-size: 17px;"">Risultato: " & Punti1 & "-" & Punti2 & "</span>"
									Stringozza &= "</td>"
									Stringozza &= "<td style=""text-align: center;"" colspan=""3"">"
									Stringozza &= "<span class=""testo "" style=""color: " & Colore & "; font-size: 15px;"">" & Esito & "</span>"
									Stringozza &= "</td>"
									Stringozza &= "</tr>"

									Dim Notelle As String = ""

									Sql = "Select * From Risultati Where idPartita = " & Rec("idPartita").Value
									Try
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Not Rec2.Eof Then
												Notelle = Rec2(2).Value
											End If
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & " " & ex.Message
									End Try

									Dim Indirizzo As String = ""

									If Rec("Casa").Value = "S" Then
										Indirizzo = Rec("IndirizzoCasa").Value
										Stringozza &= "<tr " & RitornaColoreSfondo() & ">"
										Stringozza &= "<td colspan=""6"" style=""text-align: center;""><span class=""testo nero"" style=""font-size: 14px; font-style: italic; color: #8e2e40;"">In casa. " & Rec("IndirizzoCasa").Value & "</span></td>"
										Stringozza &= "<td colspan=""5"" style=""text-align: center;""><span class=""testo"" style=""font-size: 14px; font-style: italic; color: crimson; font-variant: all-petite-caps;"">" & Notelle & "</span></td>"
										Stringozza &= "</tr>"
									Else
										If Rec("Casa").Value = "N" Then
											Indirizzo = Rec("IndirizzoCampo").Value
											Stringozza &= "<tr " & RitornaColoreSfondo() & ">"
											Stringozza &= "<td colspan=""6"" style=""text-align: center;""><span class=""testo"" style=""font-size: 14px; font-style: italic; color: #8e2e40;"">Campo " & Rec("Campo").Value & ". " & Rec("IndirizzoCampo").Value & "</span></td>"
											Stringozza &= "<td colspan=""5"" style=""text-align: center;""><span class=""testo"" style=""font-size: 14px; font-style: italic; color: crimson; font-variant: all-petite-caps;"">" & Notelle & "</span></td>"
											Stringozza &= "</tr>"
										Else
											Dim CampoEsterno As String = ""

											Sql = "Select * From CampiEsterni Where idPartita = " & Rec("idPartita").Value
											Try
												Rec2 = LeggeQuery(Conn, Sql, Connessione)
												If TypeOf (Rec2) Is String Then
													Ritorno = Rec2
												Else
													If Not Rec2.Eof Then
														CampoEsterno = Rec2(1).Value
													End If
												End If
											Catch ex As Exception
												Ritorno = StringaErrore & " " & ex.Message
											End Try

											Indirizzo = CampoEsterno
											Stringozza &= "<tr " & RitornaColoreSfondo() & ">"
											Stringozza &= "<td colspan=""6""><span class=""testo nero"" style=""font-size: 14px; font-style: italic; color: #8e2e40;"">Campo esterno: " & CampoEsterno & "</span></td>"
											Stringozza &= "<td colspan=""5""><span class=""testo"" style=""font-size: 14px; font-style: italic; color: crimson; font-variant: all-petite-caps;"">" & Notelle & "</span></td>"
											Stringozza &= "</tr>"
										End If
									End If

									' Stringozza &= "<tr>"
									' Stringozza &= "<td colspan=""11"" style=""background-color: #555; height: 5px;""></td>"
									' Stringozza &= "</tr>"

									Sql = "Select * From CoordinatePartite " &
										"Where idPartita = " & Rec("idPartita").Value
									Try
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Not Rec2.Eof Then
												If "" & Rec2("Lat").Value <> "" And "" & Rec2("Lon").Value <> "" Then
													Dim Descrizione As String = Rec("DataOra").Value & ": " & Casa & "-" & Fuori & " " & Punti1 & "-" & Punti2 & "<br />" & Esito & " - " & Tempo.Replace("<br />", " ") & "<br />"
													Dim Ok As Boolean = False

													For ii As Integer = 0 To lmLat.Count - 1
														If lmLat(ii).Trim = Rec2("Lat").Value.ToString.Trim And lmLon(ii).Trim = Rec2("Lon").Value.ToString.Trim Then
															lmDescr.Item(ii) &= "<hr />" & Descrizione

															Ok = True
															Exit For
														End If
														ii += 1
													Next

													If Not Ok Then
														lmLat.Add(Rec2("Lat").Value.ToString.Trim)
														lmLon.Add(Rec2("Lon").Value.ToString.Trim)
														lmInd.Add(Indirizzo)
														lmDescr.Add(Descrizione)
													End If
												End If
											End If
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & " " & ex.Message
									End Try

									Stringozza &= "</table>"
									Stringozza &= "</td>"
									Stringozza &= "</tr>"
									Stringozza &= "</table>"

									Rec.MoveNext
								Loop

								Select Case i
									Case 1
										ListaPartiteDiCampionato = Stringozza
									Case 2
										ListaPartiteAmichevoli = Stringozza
									Case 3
										ListaPartiteTornei = Stringozza
								End Select

								Rec.Close()
							End If
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				Next i

				For ii As Integer = 0 To lmLat.Count - 1
					ListaMarkers &= "AggiungeMarker(" & lmLat.Item(ii) & ", " & lmLon.Item(ii) & ", '" & lmInd.Item(ii) & "', '" & lmDescr.Item(ii) & "');" & vbCrLf
				Next

				If GraficoPuntiCampionato.Length > 0 Then
					GraficoPuntiCampionato = Mid(GraficoPuntiCampionato, 1, GraficoPuntiCampionato.Length - 1)
				End If

				Sql = "SELECT Tempo1Tempo, Tempo2Tempo, Tempo3Tempo " &
					"FROM RisultatiAggiuntivi INNER JOIN Partite ON RisultatiAggiuntivi.idPartita = Partite.idPartita " &
					"WHERE Partite.idAnno=" & idAnno & " AND Partite.idCategoria=" & idCategoria
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Dim SecondiTotali As Integer = 0
						Dim MinutiTotali As Integer = 0
						Dim OreTotali As Integer = 0

						Do Until Rec.Eof
							If "" & Rec("Tempo1Tempo").Value <> "" Then
								Dim t() As String = Rec("Tempo1Tempo").Value.split(":")
								Dim Minuti As Integer = Val(t(0))
								Dim Secondi As Integer = Val(t(1))

								SecondiTotali += Secondi
								MinutiTotali += Minuti
								While SecondiTotali > 60
									SecondiTotali -= 60
									MinutiTotali += 1
								End While
								While MinutiTotali > 60
									OreTotali += 1
									MinutiTotali -= 60
								End While
							End If
							If "" & Rec("Tempo2Tempo").Value <> "" Then
								Dim t() As String = Rec("Tempo2Tempo").Value.split(":")
								Dim Minuti As Integer = Val(t(0))
								Dim Secondi As Integer = Val(t(1))

								SecondiTotali += Secondi
								MinutiTotali += Minuti
								While SecondiTotali > 60
									SecondiTotali -= 60
									MinutiTotali += 1
								End While
								While MinutiTotali > 60
									OreTotali += 1
									MinutiTotali -= 60
								End While
							End If
							If "" & Rec("Tempo3Tempo").Value <> "" Then
								Dim t() As String = Rec("Tempo3Tempo").Value.split(":")
								Dim Minuti As Integer = Val(t(0))
								Dim Secondi As Integer = Val(t(1))

								SecondiTotali += Secondi
								MinutiTotali += Minuti
								While SecondiTotali > 60
									SecondiTotali -= 60
									MinutiTotali += 1
								End While
								While MinutiTotali > 60
									OreTotali += 1
									MinutiTotali -= 60
								End While
							End If

							Rec.MoveNext
						Loop
						Rec.Close()

						TempoTotaleDiGioco = Format(OreTotali, "00") & ":" & Format(MinutiTotali, "00") & ":" & Format(SecondiTotali, "00")
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				For i As Integer = 1 To 3
					Dim ListaPartite As New List(Of Integer)

					Select Case i
						Case 1
							ListaPartite = PartiteCampionato
						Case 2
							ListaPartite = PartiteAmichevoli
						Case 3
							ListaPartite = PartiteTornei
					End Select

					For Each Partita As Integer In ListaPartite
						Sql = "Select (Select RisultatoATempi From Partite Where idPartita=" & Partita & ") As RisultatoATempi, " &
							"(Select Casa From Partite Where idPartita=" & Partita & ") As Casa, *, " &
							"(Select RisGiochetti From RisultatiAggiuntivi Where idPartita =" & Partita & ") " &
							"From (" &
							"Select Sum(Goal1Tempo) As G1Tempo, Sum(Goal2Tempo) As G2Tempo, Sum(Goal3Tempo) As G3Tempo, " &
							"Sum(GA1Tempo) As GoalAvv1Tempo, Sum(Ga2Tempo) As GoalAvv2Tempo, Sum(GA3Tempo) As GoalAvv3Tempo, " &
							"(Select RisGiochetti From RisultatiAggiuntivi Where idPartita =" & Partita & ")  As RisGiochetti From (" &
							"Select 0 As Goal1Tempo, 0 As Goal2Tempo, 0 As Goal3Tempo, " &
							"IIf(GoalAvvPrimoTempo > 0, GoalAvvPrimoTempo, 0) As GA1Tempo, " &
							"IIf(GoalAvvSecondoTempo > 0, GoalAvvSecondoTempo, 0) As GA2Tempo, " &
							"IIf(GoalAvvTerzoTempo > 0, GoalAvvTerzoTempo, 0) As GA3Tempo, " &
							"RisultatiAggiuntivi.RisGiochetti " &
							"From Partite Left Join RisultatiAggiuntivi On Partite.idPartita = RisultatiAggiuntivi.idPartita " &
							"Where Partite.idPartita = " & Partita & " " &
							"Union All " &
							"Select Count(*) As Goal1Tempo, 0 As Goal2Tempo, 0 As Goal3Tempo, 0 As GA1Tempo, 0 As GA2Tempo, 0 As GA3Tempo, '' As RisGiochetti From RisultatiAggiuntiviMarcatori Where idPartita = " & Partita & " And idTempo=1 " &
							"Union All " &
							"Select 0 As Goal1Tempo, Count(*) As Goal2Tempo, 0 As Goal3Tempo, 0 As GA1Tempo, 0 As GA2Tempo, 0 As GA3Tempo, '' As RisGiochetti From RisultatiAggiuntiviMarcatori Where idPartita = " & Partita & " And idTempo=2 " &
							"Union All " &
							"Select 0 As Goal1Tempo, 0 As Goal2Tempo, Count(*) As Goal3Tempo, 0 As GA1Tempo, 0 As GA2Tempo, 0 As GA3Tempo, '' As RisGiochetti From RisultatiAggiuntiviMarcatori Where idPartita = " & Partita & " And idTempo=3 " &
							") As A) As B"
						Try
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Not Rec.Eof Then
									Dim SommaGoal As Integer = Rec("G1Tempo").Value + Rec("G2Tempo").Value + Rec("G3Tempo").Value
									Dim SommaGoalAvv As Integer = Rec("GoalAvv1Tempo").Value + Rec("GoalAvv2Tempo").Value + Rec("GoalAvv3Tempo").Value
									Dim SommaTotale As Integer = SommaGoal + SommaGoalAvv

									If SommaTotale > maxGoalInUnaPartita Then
										maxGoalInUnaPartita = SommaTotale
										PartitaConPiuGoal = Partita
									End If

									If SommaTotale < minGoalInUnaPartita Then
										minGoalInUnaPartita = SommaTotale
										PartitaConMenoGoal = Partita
									End If

									Dim RigoriPropri As Integer = 0
									Dim RigoriAvversari As Integer = 0

									Sql = "Select Count(*) As Quanti From RigoriPropri Where idPartita=" & Partita & " And Termine=1"
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											RigoriPropri = Rec2("Quanti").Value
										End If
									End If

									Sql = "Select Segnati As Quanti From RigoriAvversari Where idPartita=" & Partita
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											RigoriAvversari = Rec2("Quanti").Value
										End If
									End If

									If Rec("RisultatoATempi").Value = "N" Then
										Dim GoalTotaliFatti As Integer = Rec("G1Tempo").Value + Rec("G2Tempo").Value + Rec("G3Tempo").Value
										GoalTotaliFatti += RigoriPropri
										Dim GoalTotaliSubiti As Integer = Rec("GoalAvv1Tempo").Value + Rec("GoalAvv2Tempo").Value + Rec("GoalAvv3Tempo").Value
										GoalTotaliSubiti += RigoriAvversari

										If GoalTotaliFatti > GoalTotaliSubiti Then
											Select Case Rec("Casa").Value
												Case "S"
													Select Case i
														Case 1
															VittorieCampionatoCasa += 1
														Case 2
															VittorieAmichevoliCasa += 1
														Case 3
															VittorieTorneiCasa += 1
													End Select
												Case "N"
													Select Case i
														Case 1
															VittorieCampionatoFuori += 1
														Case 2
															VittorieAmichevoliFuori += 1
														Case 3
															VittorieTorneiFuori += 1
													End Select
												Case "E"
													Select Case i
														Case 1
															VittorieCampionatoCampoEsterno += 1
														Case 2
															VittorieAmichevoliCampoEsterno += 1
														Case 3
															VittorieTorneiCampoEsterno += 1
													End Select
											End Select
										Else
											If GoalTotaliFatti < GoalTotaliSubiti Then
												Select Case Rec("Casa").Value
													Case "S"
														Select Case i
															Case 1
																SconfitteCampionatoCasa += 1
															Case 2
																SconfitteAmichevoliCasa += 1
															Case 3
																SconfitteTorneiCasa += 1
														End Select
													Case "N"
														Select Case i
															Case 1
																SconfitteCampionatoFuori += 1
															Case 2
																SconfitteAmichevoliFuori += 1
															Case 3
																SconfitteTorneiFuori += 1
														End Select
													Case "E"
														Select Case i
															Case 1
																SconfitteCampionatoCampoEsterno += 1
															Case 2
																SconfitteAmichevoliCampoEsterno += 1
															Case 3
																SconfitteTorneiCampoEsterno += 1
														End Select
												End Select
											Else
												Select Case Rec("Casa").Value
													Case "S"
														Select Case i
															Case 1
																PareggiCampionatoCasa += 1
															Case 2
																PareggiAmichevoliCasa += 1
															Case 3
																PareggiTorneiCasa += 1
														End Select
													Case "N"
														Select Case i
															Case 1
																PareggiCampionatoFuori += 1
															Case 2
																PareggiAmichevoliFuori += 1
															Case 3
																PareggiTorneiFuori += 1
														End Select
													Case "E"
														Select Case i
															Case 1
																PareggiCampionatoCampoEsterno += 1
															Case 2
																PareggiAmichevoliCampoEsterno += 1
															Case 3
																PareggiTorneiCampoEsterno += 1
														End Select
												End Select
											End If
										End If
									Else
										Dim Punti1 As Integer = 0
										Dim Punti2 As Integer = 0

										If Rec("G1Tempo").Value > Rec("GoalAvv1Tempo").Value Then
											Punti1 += 1
										Else
											If Rec("G1Tempo").Value < Rec("GoalAvv1Tempo").Value Then
												Punti2 += 1
											Else
												Punti1 += 1
												Punti2 += 1
											End If
										End If

										If Rec("G2Tempo").Value > Rec("GoalAvv2Tempo").Value Then
											Punti1 += 1
										Else
											If Rec("G2Tempo").Value < Rec("GoalAvv2Tempo").Value Then
												Punti2 += 1
											Else
												Punti1 += 1
												Punti2 += 1
											End If
										End If

										If Rec("G3Tempo").Value > Rec("GoalAvv3Tempo").Value Then
											Punti1 += 1
										Else
											If Rec("G3Tempo").Value < Rec("GoalAvv3Tempo").Value Then
												Punti2 += 1
											Else
												Punti1 += 1
												Punti2 += 1
											End If
										End If

										If RigoriPropri > 0 Or RigoriAvversari > 0 Then
											If RigoriPropri > RigoriAvversari Then
												Punti1 += 1
											Else
												If RigoriPropri < RigoriAvversari Then
													Punti2 += 1
												Else
													Punti1 += 1
													Punti2 += 1
												End If
											End If
										End If

										Dim RisGiochetti As String = Rec("RisGiochetti").Value

										If RisGiochetti.Contains("-") Then
											Dim r() As String = RisGiochetti.Split("-")
											Dim ris1 As Integer = Val(r(0))
											Dim ris2 As Integer = Val(r(1))

											If ris1 > ris2 Then
												Punti1 += 1
											Else
												If ris1 < ris2 Then
													Punti2 += 1
												Else
													Punti1 += 1
													Punti2 += 1
												End If
											End If
										End If

										If Punti1 > Punti2 Then
											Select Case Rec("Casa").Value
												Case "S"
													Select Case i
														Case 1
															VittorieCampionatoCasa += 1
														Case 2
															VittorieAmichevoliCasa += 1
														Case 3
															VittorieTorneiCasa += 1
													End Select
												Case "N"
													Select Case i
														Case 1
															VittorieCampionatoFuori += 1
														Case 2
															VittorieAmichevoliFuori += 1
														Case 3
															VittorieTorneiFuori += 1
													End Select
												Case "E"
													Select Case i
														Case 1
															VittorieCampionatoCampoEsterno += 1
														Case 2
															VittorieAmichevoliCampoEsterno += 1
														Case 3
															VittorieTorneiCampoEsterno += 1
													End Select
											End Select
										Else
											If Punti1 < Punti2 Then
												Select Case Rec("Casa").Value
													Case "S"
														Select Case i
															Case 1
																SconfitteCampionatoCasa += 1
															Case 2
																SconfitteAmichevoliCasa += 1
															Case 3
																SconfitteTorneiCasa += 1
														End Select
													Case "N"
														Select Case i
															Case 1
																SconfitteCampionatoFuori += 1
															Case 2
																SconfitteAmichevoliFuori += 1
															Case 3
																SconfitteTorneiFuori += 1
														End Select
													Case "E"
														Select Case i
															Case 1
																SconfitteCampionatoCampoEsterno += 1
															Case 2
																SconfitteAmichevoliCampoEsterno += 1
															Case 3
																SconfitteTorneiCampoEsterno += 1
														End Select
												End Select
											Else
												Select Case Rec("Casa").Value
													Case "S"
														Select Case i
															Case 1
																PareggiCampionatoCasa += 1
															Case 2
																PareggiAmichevoliCasa += 1
															Case 3
																PareggiTorneiCasa += 1
														End Select
													Case "N"
														Select Case i
															Case 1
																PareggiCampionatoFuori += 1
															Case 2
																PareggiAmichevoliFuori += 1
															Case 3
																PareggiTorneiFuori += 1
														End Select
													Case "E"
														Select Case i
															Case 1
																PareggiCampionatoCampoEsterno += 1
															Case 2
																PareggiAmichevoliCampoEsterno += 1
															Case 3
																PareggiTorneiCampoEsterno += 1
														End Select
												End Select
											End If
										End If
									End If
								End If

								Rec.Close()
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Exit For
						End Try
					Next
				Next

				Sql = "SELECT TipologiePartite.Descrizione As Tipologia, NomeSquadra, SquadreAvversarie.Descrizione, Partite.Casa, Partite.DataOra, Tempo, Gradi, SquadreAvversarie.idAvversario, " &
					"(Select Count(*) From RisultatiAggiuntiviMarcatori Where idPartita=" & PartitaConPiuGoal & ") As Goal, " &
					"(Select Sum(iif(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0))+Sum(iif(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0))+Sum(iif(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) " &
					"From RisultatiAggiuntivi Where idPartita=" & PartitaConPiuGoal & ") As GoalAvv From " &
					 "(((Partite Left Join SquadreAvversarie On Partite.idAvversario = SquadreAvversarie.idAvversario) " &
					"Left Join Anni On Partite.idAnno = Anni.idAnno) Left Join MeteoPartite On Partite.idPartita = MeteoPartite.idPartita) " &
					"Left Join TipologiePartite On Partite.idTipologia = TipologiePartite.idTipologia " &
					"Where Partite.idAnno = " & idAnno & " And Partite.idCategoria = " & idCategoria & " And Partite.idPartita = " & PartitaConPiuGoal & ""
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof Then
							Dim Casa As String
							Dim Fuori As String
							Dim Imm1 As String = PathBaseImmagini & "/" & Squadra & "/Categorie/" & idAnno & "_" & idCategoria & ".Jpg"
							Dim Imm2 As String = PathBaseImmagini & "/Avversari/" & Rec("idAvversario").Value & ".Jpg"
							Dim ImmCasa As String
							Dim ImmFuori As String

							If Rec("Casa").Value = "S" Then
								Casa = Rec("NomeSquadra").Value
								Fuori = Rec("Descrizione").Value
								ImmCasa = "<td style ="" border: 1px solid #999; text-align: center;""><img src=""" & Imm1 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & Pathimmagini & "Sconosciuto.png'""  /></td>"
								ImmFuori = "<td style ="" border: 1px solid #999; text-align: center;""><img src=""" & Imm2 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & PathImmagini & "Sconosciuto.png'""  /></td>"
							Else
								Fuori = Rec("NomeSquadra").Value
								Casa = Rec("Descrizione").Value
								ImmCasa = "<td style ="" border: 1px solid #999; text-align: center;""><img src=""" & Imm2 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
								ImmFuori = "<td style ="" border: 1px solid #999; text-align: center;""><img src=""" & Imm1 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
							End If

							sPartitaConPiuGoal &= "<table cellspacing=""0"" style =""width: 100%; border: 1px solid #999;"">"
							sPartitaConPiuGoal &= "<tr " & RitornaColoreSfondo() & ">"
							sPartitaConPiuGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConPiuGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("Tipologia").Value & "</span>"
							sPartitaConPiuGoal &= "</td>"
							sPartitaConPiuGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConPiuGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("DataOra").Value & "</span>"
							sPartitaConPiuGoal &= "</td>"
							sPartitaConPiuGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"

							Dim Tempo As String = ""

							If "" & Rec("Tempo").Value <> "" Then
								Tempo = Rec("Tempo").Value
							End If
							If "" & Rec("Gradi").Value <> "" Then
								Tempo &= " " & Rec("Gradi").Value & " Gradi"
							End If

							sPartitaConPiuGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Tempo & "</span>"
							sPartitaConPiuGoal &= "</td>"
							sPartitaConPiuGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConPiuGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Casa & "</span>"
							sPartitaConPiuGoal &= "</td>"
							sPartitaConPiuGoal &= ImmCasa
							sPartitaConPiuGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConPiuGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("Goal").Value & "-" & Rec("GoalAvv").Value & "</span>"
							sPartitaConPiuGoal &= "</td>"
							sPartitaConPiuGoal &= ImmFuori
							sPartitaConPiuGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConPiuGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Fuori & "</span>"
							sPartitaConPiuGoal &= "</td>"
							sPartitaConPiuGoal &= "</tr>"
							sPartitaConPiuGoal &= "</table>"

							Rec.Close()
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Sql = "SELECT TipologiePartite.Descrizione As Tipologia, NomeSquadra, SquadreAvversarie.Descrizione, Partite.Casa, Partite.DataOra, Tempo, Gradi, SquadreAvversarie.idAvversario, " &
					"(Select Count(*) From RisultatiAggiuntiviMarcatori Where idPartita=" & PartitaConMenoGoal & ") As Goal, " &
					"(Select Sum(iif(GoalAvvPrimoTempo>0,GoalAvvPrimoTempo,0))+Sum(iif(GoalAvvSecondoTempo>0,GoalAvvSecondoTempo,0))+Sum(iif(GoalAvvTerzoTempo>0,GoalAvvTerzoTempo,0)) " &
					"From RisultatiAggiuntivi Where idPartita=" & PartitaConMenoGoal & ") As GoalAvv From " &
					"(((Partite Left Join SquadreAvversarie On Partite.idAvversario = SquadreAvversarie.idAvversario) " &
					"Left Join Anni On Partite.idAnno = Anni.idAnno) Left Join MeteoPartite On Partite.idPartita = MeteoPartite.idPartita) " &
					"Left Join TipologiePartite On Partite.idTipologia = TipologiePartite.idTipologia " &
					"Where Partite.idAnno = " & idAnno & " And Partite.idCategoria = " & idCategoria & " And Partite.idPartita = " & PartitaConMenoGoal & ""
				Try
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof Then
							Dim Casa As String
							Dim Fuori As String
							Dim Imm1 As String = PathBaseImmagini & "/" & Squadra & "/Categorie/" & idAnno & "_" & idCategoria & ".Jpg"
							Dim Imm2 As String = PathBaseImmagini & "/Avversari/" & Rec("idAvversario").Value & ".Jpg"
							Dim ImmCasa As String
							Dim ImmFuori As String

							If Rec("Casa").Value = "S" Then
								Casa = Rec("NomeSquadra").Value
								Fuori = Rec("Descrizione").Value
								ImmCasa = "<td style ="" border: 1px solid #999; text-align: center;""><img src=""" & Imm1 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
								ImmFuori = "<td style ="" border: 1px solid #999; text-align: center;""><img src=""" & Imm2 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
							Else
								Fuori = Rec("NomeSquadra").Value
								Casa = Rec("Descrizione").Value
								ImmCasa = "<td style ="" border: 1px solid #999;""><img src=""" & Imm2 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
								ImmFuori = "<td style ="" border: 1px solid #999;""><img src=""" & Imm1 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
							End If

							sPartitaConMenoGoal &= "<table cellspacing=""0"" style =""width: 100%; border: 1px solid #999;"">"
							sPartitaConMenoGoal &= "<tr " & RitornaColoreSfondo() & ">"
							sPartitaConMenoGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConMenoGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("Tipologia").Value & "</span>"
							sPartitaConMenoGoal &= "</td>"
							sPartitaConMenoGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConMenoGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("DataOra").Value & "</span>"
							sPartitaConMenoGoal &= "</td>"
							sPartitaConMenoGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"

							Dim Tempo As String = ""

							If "" & Rec("Tempo").Value <> "" Then
								Tempo = Rec("Tempo").Value
							End If
							If "" & Rec("Gradi").Value <> "" Then
								Tempo &= " " & Rec("Gradi").Value & " Gradi"
							End If

							sPartitaConMenoGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Tempo & "</span>"
							sPartitaConMenoGoal &= "</td>"
							sPartitaConMenoGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConMenoGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Casa & "</span>"
							sPartitaConMenoGoal &= "</td>"
							sPartitaConMenoGoal &= ImmCasa
							sPartitaConMenoGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConMenoGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Rec("Goal").Value & "-" & Rec("GoalAvv").Value & "</span>"
							sPartitaConMenoGoal &= "</td>"
							sPartitaConMenoGoal &= ImmFuori
							sPartitaConMenoGoal &= "<td style ="" border: 1px solid #999; text-align: center;"">"
							sPartitaConMenoGoal &= "<span class=""testo nero"" style=""font-size: 16px;"">" & Fuori & "</span>"
							sPartitaConMenoGoal &= "</td>"
							sPartitaConMenoGoal &= "</tr>"
							sPartitaConMenoGoal &= "</table>"

							Rec.Close()
						End If
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try

				Dim Filone As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_statistiche.txt")
				gf.CreaDirectoryDaPercorso(HttpContext.Current.Server.MapPath(".") & "\Statistiche\" & Squadra & "\")
				Dim NomeFileFinale As String = HttpContext.Current.Server.MapPath(".") & "\Statistiche\" & Squadra & "\" & idAnno & "_" & idCategoria & ".html"

				Filone = Filone.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")
				Filone = Filone.Replace("***ANNO***", idAnno)
				Filone = Filone.Replace("***CATEGORIA***", idCategoria)

				Dim Stringona As String = ""

				For Each Giocatore As String In NomiMarcatoriCampionatoCasa
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliCasa
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliCasa.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliCasa.Add(Giocatore)
					End If
				Next
				For Each Giocatore As String In NomiMarcatoriAmichevoliCasa
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliCasa
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString.Trim
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliCasa.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliCasa.Add(Giocatore)
					End If
				Next
				For Each Giocatore As String In NomiMarcatoriTorneiCasa
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliCasa
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliCasa.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliCasa.Add(Giocatore)
					End If
				Next

				For Each Giocatore As String In NomiMarcatoriCampionatoFuori
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliFuori
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliFuori.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliFuori.Add(Giocatore)
					End If
				Next
				For Each Giocatore As String In NomiMarcatoriAmichevoliFuori
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliFuori
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliFuori.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliFuori.Add(Giocatore)
					End If
				Next
				For Each Giocatore As String In NomiMarcatoriTorneiFuori
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliFuori
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliFuori.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliFuori.Add(Giocatore)
					End If
				Next

				For Each Giocatore As String In NomiMarcatoriCampionatoCampoEsterno
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliCampoEsterno
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliCampoEsterno.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliCampoEsterno.Add(Giocatore)
					End If
				Next
				For Each Giocatore As String In NomiMarcatoriAmichevoliCampoEsterno
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliCampoEsterno
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliCampoEsterno.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliCampoEsterno.Add(Giocatore)
					End If
				Next
				For Each Giocatore As String In NomiMarcatoriTorneiCampoEsterno
					Dim Giocatore2 As String = ""
					Dim c1() As String = Giocatore.Split("-")
					Dim i As Integer = 0
					For Each Gioc2 As String In NomiMarcatoriGeneraliCampoEsterno
						Dim c2() As String = Gioc2.Split("-")
						If c1(1) = c2(1) Then
							c2(2) = (Val(c2(2)) + Val(c1(2))).ToString
							Giocatore2 = c1(0) + "-" + c1(1) + "-" + (c2(2))
							NomiMarcatoriGeneraliCampoEsterno.Item(i) = Giocatore2
							Exit For
						End If
						i += 1
					Next
					If Giocatore2 = "" Then
						NomiMarcatoriGeneraliCampoEsterno.Add(Giocatore)
					End If
				Next

				For Each Dove As String In PartiteCampionatoDove
					Select Case Dove
						Case "S"
							GiocateCampionatoCasa += 1
						Case "N"
							GiocateCampionatoFuori += 1
						Case "E"
							GiocateCampionatoCampoEsterno += 1
					End Select
				Next

				For Each Dove As String In PartiteAmichevoliDove
					Select Case Dove
						Case "S"
							GiocateAmichevoliCasa += 1
						Case "N"
							GiocateAmichevoliFuori += 1
						Case "E"
							GiocateAmichevoliCampoEsterno += 1
					End Select
				Next

				For Each Dove As String In PartiteTorneiDove
					Select Case Dove
						Case "S"
							GiocateTorneiCasa += 1
						Case "N"
							GiocateTorneiFuori += 1
						Case "E"
							GiocateTorneiCampoEsterno += 1
					End Select
				Next

				For i As Integer = 0 To NomiMarcatoriGeneraliCasa.Count - 1
					Dim c1() As String = NomiMarcatoriGeneraliCasa.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriGeneraliCasa.Count - 1
						Dim c2() As String = NomiMarcatoriGeneraliCasa.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriGeneraliCasa.Item(i)
							NomiMarcatoriGeneraliCasa.Item(i) = NomiMarcatoriGeneraliCasa.Item(k)
							NomiMarcatoriGeneraliCasa.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriGeneraliFuori.Count - 1
					Dim c1() As String = NomiMarcatoriGeneraliFuori.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriGeneraliFuori.Count - 1
						Dim c2() As String = NomiMarcatoriGeneraliFuori.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriGeneraliFuori.Item(i)
							NomiMarcatoriGeneraliFuori.Item(i) = NomiMarcatoriGeneraliFuori.Item(k)
							NomiMarcatoriGeneraliFuori.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriGeneraliCampoEsterno.Count - 1
					Dim c1() As String = NomiMarcatoriGeneraliCampoEsterno.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriGeneraliCampoEsterno.Count - 1
						Dim c2() As String = NomiMarcatoriGeneraliCampoEsterno.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriGeneraliCampoEsterno.Item(i)
							NomiMarcatoriGeneraliCampoEsterno.Item(i) = NomiMarcatoriGeneraliCampoEsterno.Item(k)
							NomiMarcatoriGeneraliCampoEsterno.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriCampionatoCasa.Count - 1
					Dim c1() As String = NomiMarcatoriCampionatoCasa.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriCampionatoCasa.Count - 1
						Dim c2() As String = NomiMarcatoriCampionatoCasa.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriCampionatoCasa.Item(i)
							NomiMarcatoriCampionatoCasa.Item(i) = NomiMarcatoriCampionatoCasa.Item(k)
							NomiMarcatoriCampionatoCasa.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriCampionatoFuori.Count - 1
					Dim c1() As String = NomiMarcatoriCampionatoFuori.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriCampionatoFuori.Count - 1
						Dim c2() As String = NomiMarcatoriCampionatoFuori.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriCampionatoFuori.Item(i)
							NomiMarcatoriCampionatoFuori.Item(i) = NomiMarcatoriCampionatoFuori.Item(k)
							NomiMarcatoriCampionatoFuori.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriCampionatoCampoEsterno.Count - 1
					Dim c1() As String = NomiMarcatoriCampionatoCampoEsterno.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriCampionatoCampoEsterno.Count - 1
						Dim c2() As String = NomiMarcatoriCampionatoCampoEsterno.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriCampionatoCampoEsterno.Item(i)
							NomiMarcatoriCampionatoCampoEsterno.Item(i) = NomiMarcatoriCampionatoCampoEsterno.Item(k)
							NomiMarcatoriCampionatoCampoEsterno.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriAmichevoliCasa.Count - 1
					Dim c1() As String = NomiMarcatoriAmichevoliCasa.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriAmichevoliCasa.Count - 1
						Dim c2() As String = NomiMarcatoriAmichevoliCasa.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriAmichevoliCasa.Item(i)
							NomiMarcatoriAmichevoliCasa.Item(i) = NomiMarcatoriAmichevoliCasa.Item(k)
							NomiMarcatoriAmichevoliCasa.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriAmichevoliFuori.Count - 1
					Dim c1() As String = NomiMarcatoriAmichevoliFuori.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriAmichevoliFuori.Count - 1
						Dim c2() As String = NomiMarcatoriAmichevoliFuori.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriAmichevoliFuori.Item(i)
							NomiMarcatoriAmichevoliFuori.Item(i) = NomiMarcatoriAmichevoliFuori.Item(k)
							NomiMarcatoriAmichevoliFuori.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriAmichevoliCampoEsterno.Count - 1
					Dim c1() As String = NomiMarcatoriAmichevoliCampoEsterno.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriAmichevoliCampoEsterno.Count - 1
						Dim c2() As String = NomiMarcatoriAmichevoliCampoEsterno.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriAmichevoliCampoEsterno.Item(i)
							NomiMarcatoriAmichevoliCampoEsterno.Item(i) = NomiMarcatoriAmichevoliCampoEsterno.Item(k)
							NomiMarcatoriAmichevoliCampoEsterno.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriTorneiCasa.Count - 1
					Dim c1() As String = NomiMarcatoriTorneiCasa.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriTorneiCasa.Count - 1
						Dim c2() As String = NomiMarcatoriTorneiCasa.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriTorneiCasa.Item(i)
							NomiMarcatoriTorneiCasa.Item(i) = NomiMarcatoriTorneiCasa.Item(k)
							NomiMarcatoriTorneiCasa.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriTorneiFuori.Count - 1
					Dim c1() As String = NomiMarcatoriTorneiFuori.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriTorneiFuori.Count - 1
						Dim c2() As String = NomiMarcatoriTorneiFuori.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriTorneiFuori.Item(i)
							NomiMarcatoriTorneiFuori.Item(i) = NomiMarcatoriTorneiFuori.Item(k)
							NomiMarcatoriTorneiFuori.Item(k) = Appoggio
						End If
					Next
				Next

				For i As Integer = 0 To NomiMarcatoriTorneiCampoEsterno.Count - 1
					Dim c1() As String = NomiMarcatoriTorneiCampoEsterno.Item(i).Split("-")
					For k As Integer = 0 To NomiMarcatoriTorneiCampoEsterno.Count - 1
						Dim c2() As String = NomiMarcatoriTorneiCampoEsterno.Item(k).Split("-")
						If Val(c1(2) > Val(c2(2))) Then
							Dim Appoggio As String = NomiMarcatoriTorneiCampoEsterno.Item(i)
							NomiMarcatoriTorneiCampoEsterno.Item(i) = NomiMarcatoriTorneiCampoEsterno.Item(k)
							NomiMarcatoriTorneiCampoEsterno.Item(k) = Appoggio
						End If
					Next
				Next

				' Generali casa
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateCampionatoCasa + GiocateAmichevoliCasa + GiocateTorneiCasa).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieCampionatoCasa + VittorieAmichevoliCampoEsterno + VittorieAmichevoliCasa).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiCampionatoCasa + PareggiAmichevoliCasa + PareggiTorneiCasa).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteCampionatoCasa + SconfitteAmichevoliCasa + SconfitteTorneiCasa).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalCampionatoCasa + GoalAmichevoliCasa + GoalTorneiCasa).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvCampionatoCasa1Tempo + GoalAvvCampionatoCasa2Tempo + GoalAvvCampionatoCasa3Tempo +
				GoalAvvAmichevoliCasa1Tempo + GoalAvvAmichevoliCasa2Tempo + GoalAvvAmichevoliCasa3Tempo +
				GoalAvvTorneiCasa1Tempo + GoalAvvTorneiCasa2Tempo + GoalAvvTorneiCasa3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_GENERALI_CASA***", Stringona)

				Filone = Filone.Replace("***TEMPO_TOTALE_DI_GIOCO***", TempoTotaleDiGioco)

				' Marcatori generali casa
				Stringona = "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriGeneraliCasa
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg.ToString & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_GENERALI_MARCATORI_CASA***", Stringona)

				' Generali fuori
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateCampionatoFuori + GiocateAmichevoliFuori + GiocateTorneiFuori).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieCampionatoFuori + VittorieAmichevoliCampoEsterno + VittorieAmichevoliFuori).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiCampionatoFuori + PareggiAmichevoliFuori + PareggiTorneiFuori).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteCampionatoFuori + SconfitteAmichevoliFuori + SconfitteTorneiFuori).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalCampionatoFuori + GoalAmichevoliFuori + GoalTorneiFuori).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvCampionatoFuori1Tempo + GoalAvvCampionatoFuori2Tempo + GoalAvvCampionatoFuori3Tempo +
				GoalAvvAmichevoliFuori1Tempo + GoalAvvAmichevoliFuori2Tempo + GoalAvvAmichevoliFuori3Tempo +
				GoalAvvTorneiFuori1Tempo + GoalAvvTorneiFuori2Tempo + GoalAvvTorneiFuori3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_GENERALI_FUORI***", Stringona)

				' Marcatori generali fuori
				Stringona = ""
				Stringona = "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriGeneraliFuori
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_GENERALI_MARCATORI_FUORI***", Stringona)

				' Generali campo esterno
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateCampionatoCampoEsterno + GiocateAmichevoliCampoEsterno + GiocateTorneiCampoEsterno).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieCampionatoCampoEsterno + VittorieAmichevoliCampoEsterno + VittorieAmichevoliCampoEsterno).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiCampionatoCampoEsterno + PareggiAmichevoliCampoEsterno + PareggiTorneiCampoEsterno).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteCampionatoCampoEsterno + SconfitteAmichevoliCampoEsterno + SconfitteTorneiCampoEsterno).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalCampionatoCampoEsterno + GoalAmichevoliCampoEsterno + GoalTorneiCampoEsterno).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvCampionatoCampoEsterno1Tempo + GoalAvvCampionatoCampoEsterno2Tempo + GoalAvvCampionatoCampoEsterno3Tempo +
				GoalAvvAmichevoliCampoEsterno1Tempo + GoalAvvAmichevoliCampoEsterno2Tempo + GoalAvvAmichevoliCampoEsterno3Tempo +
				GoalAvvTorneiCampoEsterno1Tempo + GoalAvvTorneiCampoEsterno2Tempo + GoalAvvTorneiCampoEsterno3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_GENERALI_CAMPOESTERNO***", Stringona)

				' Marcatori generali campo esterno
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriGeneraliCampoEsterno
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_GENERALI_MARCATORI_CAMPOESTERNO***", Stringona)

				' Campionato casa
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateCampionatoCasa).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieCampionatoCasa).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiCampionatoCasa).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteCampionatoCasa).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalCampionatoCasa).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvCampionatoCasa1Tempo + GoalAvvCampionatoCasa2Tempo + GoalAvvCampionatoCasa3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_CAMPIONATO_CASA***", Stringona)

				' Marcatori campionato casa
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriCampionatoCasa
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_CAMPIONATO_MARCATORI_CASA***", Stringona)

				' Campionato fuori
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateCampionatoFuori).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieCampionatoFuori).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiCampionatoFuori).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteCampionatoFuori).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalCampionatoFuori).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvCampionatoFuori1Tempo + GoalAvvCampionatoFuori2Tempo + GoalAvvCampionatoFuori3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_CAMPIONATO_FUORI***", Stringona)

				' Marcatori campionato fuori
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriCampionatoFuori
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_CAMPIONATO_MARCATORI_FUORI***", Stringona)

				' Campionato campo esterno
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateCampionatoCampoEsterno).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieCampionatoCampoEsterno).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiCampionatoCampoEsterno).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteCampionatoCampoEsterno).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalCampionatoCampoEsterno).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvCampionatoCampoEsterno1Tempo + GoalAvvCampionatoCampoEsterno2Tempo + GoalAvvCampionatoCampoEsterno3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_CAMPIONATO_CAMPOESTERNO***", Stringona)

				' Marcatori campionato campo esterno
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriCampionatoCampoEsterno
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_CAMPIONATO_MARCATORI_CAMPOESTERNO***", Stringona)

				' Amichevoli casa
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateAmichevoliCasa).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieAmichevoliCasa).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiAmichevoliCasa).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteAmichevoliCasa).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalAmichevoliCasa).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvAmichevoliCasa1Tempo + GoalAvvAmichevoliCasa2Tempo + GoalAvvAmichevoliCasa3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_AMICHEVOLI_CASA***", Stringona)

				' Marcatori amichevoli casa
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriAmichevoliCasa
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_AMICHEVOLI_MARCATORI_CASA***", Stringona)

				' Amichevoli fuori
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateAmichevoliFuori).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieAmichevoliFuori).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiAmichevoliFuori).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteAmichevoliFuori).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalAmichevoliFuori).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvAmichevoliFuori1Tempo + GoalAvvAmichevoliFuori2Tempo + GoalAvvAmichevoliFuori3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_AMICHEVOLI_FUORI***", Stringona)

				' Marcatori amichevoli fuori
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriAmichevoliFuori
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_AMICHEVOLI_MARCATORI_FUORI***", Stringona)

				' Amichevoli campo esterno
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateAmichevoliCampoEsterno).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieAmichevoliCampoEsterno).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiAmichevoliCampoEsterno).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteAmichevoliCampoEsterno).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalAmichevoliCampoEsterno).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvAmichevoliCampoEsterno1Tempo + GoalAvvAmichevoliCampoEsterno2Tempo + GoalAvvAmichevoliCampoEsterno3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_AMICHEVOLI_CAMPOESTERNO***", Stringona)

				' Marcatori amichevoli campo esterno
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriAmichevoliCampoEsterno
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_AMICHEVOLI_MARCATORI_CAMPOESTERNO***", Stringona)

				' Tornei casa
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateTorneiCasa).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieTorneiCasa).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiTorneiCasa).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteTorneiCasa).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalTorneiCasa).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvTorneiCasa1Tempo + GoalAvvTorneiCasa2Tempo + GoalAvvTorneiCasa3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_TORNEI_CASA***", Stringona)

				' Marcatori Tornei casa
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriTorneiCasa
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_TORNEI_MARCATORI_CASA***", Stringona)

				' Tornei fuori
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateTorneiFuori).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieTorneiFuori).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiTorneiFuori).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteTorneiFuori).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalTorneiFuori).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvTorneiFuori1Tempo + GoalAvvTorneiFuori2Tempo + GoalAvvTorneiFuori3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_TORNEI_FUORI***", Stringona)

				' Marcatori Tornei fuori
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriTorneiFuori
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_TORNEI_MARCATORI_FUORI***", Stringona)

				' Tornei campo esterno
				Stringona = ""
				Stringona &= "Giocate: " & (GiocateTorneiCampoEsterno).ToString & "<hr />"

				Stringona &= "Vittorie: " & (VittorieTorneiCampoEsterno).ToString & "<br />"
				Stringona &= "Pareggi: " & (PareggiTorneiCampoEsterno).ToString & "<br />"
				Stringona &= "Sconfitte: " & (SconfitteTorneiCampoEsterno).ToString & "<hr />"

				Stringona &= "Goal segnati: " & (GoalTorneiCampoEsterno).ToString & "<br />"
				Stringona &= "Goal subiti: " & (GoalAvvTorneiCampoEsterno1Tempo + GoalAvvTorneiCampoEsterno2Tempo + GoalAvvTorneiCampoEsterno3Tempo).ToString & "<br />"
				Filone = Filone.Replace("***DATI_TORNEI_CAMPOESTERNO***", Stringona)

				' Marcatori amichevoli campo esterno
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In NomiMarcatoriTorneiCampoEsterno
					Dim c() As String = Giocatore.Split("-")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & c(0) & ".jpg"
					Dim gg As String = c(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td>" & gg & "</td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'"" /></td>"
					Stringona &= "<td>" & c(1) & "</td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***DATI_TORNEI_MARCATORI_CAMPOESTERNO***", Stringona)

				' MARCATORI GENERALI
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In MarcatoriGenerali
					Dim s() As String = Giocatore.Split(";")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & s(0) & ".jpg"
					Dim gg As String = s(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & gg & "</span></td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & s(1) & "</span></td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***MARCATORI_GLOBALI***", Stringona)

				' PRESENZE
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In Presenze
					Dim s() As String = Giocatore.Split(";")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & s(0) & ".jpg"
					Dim gg As String = s(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & gg & "</span></td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & s(1) & "</span></td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***PRESENZE***", Stringona)

				' SQUADRE INCONTRATE
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Squadra2 As String In SquadreIncontrate
					Dim s() As String = Squadra2.Split(";")
					Dim Imm2 As String = PathBaseImmagini & "/Avversari/" & s(0) & ".Jpg"
					Dim gg As String = s(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & gg & "</span></td>"
					Stringona &= "<td><img src=""" & Imm2 & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & s(1) & "</span></td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***SQUADRE_INCONTRATE***", Stringona)


				' MARCATORI TUTTE LE CATEGORIE
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In MarcatoriTutte
					Dim s() As String = Giocatore.Split(";")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & s(0) & ".jpg"
					Dim gg As String = s(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & gg & "</span></td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & s(1) & "</span></td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***MARCATORI_TUTTE***", Stringona)

				' PRESENZE TUTTE
				Stringona = ""
				Stringona &= "<table cellspacing=""0"" style=""width: 100%;"">"
				For Each Giocatore As String In PresenzeTutte
					Dim s() As String = Giocatore.Split(";")
					Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & s(0) & ".jpg"
					Dim gg As String = s(2).Trim
					If gg.Length = 1 Then gg = "&nbsp;" & gg

					Stringona &= "<tr " & RitornaColoreSfondo() & ">"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & gg & "</span></td>"
					Stringona &= "<td><img src=""" & Path & """ style=""width: 60px; height: 60px;"" onerror=""this.src='" & pathImmagini & "Sconosciuto.png'""  /></td>"
					Stringona &= "<td><span class=""testo nero"" style=""font-size: 16px;"">" & s(1) & "</span></td>"
					Stringona &= "</tr>"
				Next
				Stringona &= "</table>"
				Filone = Filone.Replace("***PRESENZE_TUTTE***", Stringona)

				Filone = Filone.Replace("***PARTITA_CON_PIU_GOAL***", sPartitaConPiuGoal)
				Filone = Filone.Replace("***PARTITA_CON_MENO_GOAL***", sPartitaConMenoGoal)

				Filone = Filone.Replace("***LISTA_PARTITE_CAMPIONATO***", ListaPartiteDiCampionato)
				Filone = Filone.Replace("***LISTA_PARTITE_AMICHEVOLI***", ListaPartiteAmichevoli)
				Filone = Filone.Replace("***LISTA_PARTITE_TORNEI***", ListaPartiteTornei)

				Filone = Filone.Replace("***GOALS_PER_MINUTO***", GoalFattiPerMinuto)
				Filone = Filone.Replace("***SUBITI_PER_MINUTO***", GoalSubitiPerMinuto)

				Filone = Filone.Replace("***LAT_CASA***", LatCasa)
				Filone = Filone.Replace("***LON_CASA***", LonCasa)

				Filone = Filone.Replace("***GRAFICO_PUNTI***", GraficoPuntiCampionato)
				Filone = Filone.Replace("***LISTA_MARKERS_MAPPA***", ListaMarkers)

				Filone = Filone.Replace("**CATEGORIA***", idCategoria)
				Filone = Filone.Replace("***IMM_CATEGORIA***", idAnno.ToString.Trim & "_" & idCategoria.ToString.Trim)
				Filone = Filone.Replace("***DESC_ANNO***", descAnno)
				Filone = Filone.Replace("***NOME_SQUADRA***", nomeSquadra)

				gf.CreaAggiornaFile(NomeFileFinale, Filone)

				If Ritorno = "" Then Ritorno = "OK"
			End If
		End If

		If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

		Return Ritorno
	End Function

	Private Function RitornaColoreSfondo() As String
        RigaPari = Not RigaPari
        If RigaPari Then
            Return "style=""background-color: #ccc;"""
        Else
            Return "style=""background-color: #aaa;"""
        End If
    End Function
End Class