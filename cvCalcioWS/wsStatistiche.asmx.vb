Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_stat.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsStatistiche
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function RitornaStatisticheAvversari(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaStatisticheConvocati(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaStatisticheMarcatori(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
                    Sql &= "WHERE Partite.idAnno=" & idAnno & " And Partite.Giocata='S' And Cognome Is Not Null And Categorie.idCategoria=" & idCategoria & " "
                Else
                    Sql &= "WHERE Partite.Giocata='S' And Cognome Is Not Null And Categorie.idCategoria=" & idCategoria & " "
                End If
                Sql &= "GROUP BY Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Giocatori.NumeroMaglia "
                Sql &= ") AS q "
                Sql &= "Group BY q.idGiocatore, q.Cognome, q.Nome, q.NumeroMaglia "
                Sql &= "ORDER BY 4 DESC, 2, 3"

                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Ritorno &= Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("GoalFinali").Value & ";" & Rec("NumeroMaglia").Value & "§"

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
    Public Function RitornaStatisticheRisultati(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaStatisticheMappa(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaStatisticheMinutiGoal(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaStatisticheMeteo(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaStatisticheGoalSegnatiSubiti(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaAndamento(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
                Sql &= "Group By Partite.idPartita, Partite.Casa, 0 "
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
    Public Function RitornaTipologiePartite(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaPartiteCasaFuori(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
                Sql &="FROM Partite "
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
    Public Function RitornaPartiteAllenatore(idAnno As String, SoloAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

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
    Public Function RitornaStatisticheStagione(idAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
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

                Dim GoalCampionatoCasa As Integer = 0
                Dim GoalCampionatoFuori As Integer = 0
                Dim GoalCampionatoCampoEsterno As Integer = 0

                Dim GoalAmichevoliCasa As Integer = 0
                Dim GoalAmichevoliFuori As Integer = 0
                Dim GoalAmichevoliCampoEsterno As Integer = 0

                Dim GoalTorneiCasa As Integer = 0
                Dim GoalTorneiFuori As Integer = 0
                Dim GoalTorneiCampoEsterno As Integer = 0

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

                Dim TipologiaPartitePerAnno As String = ""

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
                End If
                Sql = "SELECT 'GoalCampionatoCasa' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' " &
                    "Union All " &
                    "SELECT 'GoalCampionatoFuori' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' " &
                    "Union All " &
                    "SELECT 'GoalCampionatoCampoEsterno' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E'"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "GoalCampionatoCasa"
                                    GoalCampionatoCasa = Rec(1).Value
                                Case "GoalCampionatoFuori"
                                    GoalCampionatoFuori = Rec(1).Value
                                Case "GoalCampionatoCasa"
                                    GoalCampionatoCampoEsterno = Rec(1).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                If PartiteTorneiIN.Length > 0 Then
                    PartiteTorneiIN = Mid(PartiteTorneiIN, 1, PartiteTorneiIN.Length - 1)
                End If
                Sql = "SELECT 'GoalTorneiCasa' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' " &
                    "Union All " &
                    "SELECT 'GoalTorneiFuori' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' " &
                    "Union All " &
                    "SELECT 'GoalTorneiCampoEsterno' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E'"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "GoalTorneiCasa"
                                    GoalTorneiCasa = Rec(1).Value
                                Case "GoalTorneiFuori"
                                    GoalTorneiFuori = Rec(1).Value
                                Case "GoalTorneiCampoEsterno"
                                    GoalTorneiCampoEsterno = Rec(1).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                If PartiteAmichevoliIN.Length > 0 Then
                    PartiteAmichevoliIN = Mid(PartiteAmichevoliIN, 1, PartiteAmichevoliIN.Length - 1)
                End If
                Sql = "SELECT 'GoalAmichevoliCasa' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' " &
                    "Union All " &
                    "SELECT 'GoalAmichevoliFuori' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' " &
                    "Union All " &
                    "SELECT 'GoalAmichevoliCampoEsterno' As Cosa, Count(*) As GoalTotali " &
                    "From RisultatiAggiuntiviMarcatori Left Join Partite On RisultatiAggiuntiviMarcatori.idPartita=Partite.idPartita " &
                    "Where Partite.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E'"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "GoalAmichevoliCasa"
                                    GoalAmichevoliCasa = Rec(1).Value
                                Case "GoalAmichevoliFuori"
                                    GoalAmichevoliFuori = Rec(1).Value
                                Case "GoalAmichevoliCampoEsterno"
                                    GoalAmichevoliCampoEsterno = Rec(1).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Sql = "SELECT 'MarcatoriCasaCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome " &
                    "Union All " &
                    "Select 'MarcatoriFuoriCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome " &
                    "Union All " &
                    "Select 'MarcatoriCampoEsternoCampionato' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "MarcatoriCasaCampionato"
                                    NomiMarcatoriCampionatoCasa.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriCampionatoCasa = Rec(3).Value
                                Case "MarcatoriFuoriCampionato"
                                    NomiMarcatoriCampionatoFuori.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriCampionatoFuori = Rec(3).Value
                                Case "MarcatoriCampoEsternoCampionato"
                                    NomiMarcatoriCampionatoCampoEsterno.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriCampionatoCampoEsterno = Rec(3).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Sql = "SELECT 'MarcatoriCasaAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome " &
                    "Union All " &
                    "Select 'MarcatoriFuoriAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome " &
                    "Union All " &
                    "Select 'MarcatoriCampoEsternoAmichevoli' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "MarcatoriCasaAmichevoli"
                                    NomiMarcatoriAmichevoliCasa.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriAmichevoliCasa = Rec(3).Value
                                Case "MarcatoriFuoriAmichevoli"
                                    NomiMarcatoriAmichevoliFuori.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriAmichevoliFuori = Rec(3).Value
                                Case "MarcatoriCampoEsternoAmichevoli"
                                    NomiMarcatoriAmichevoliCampoEsterno.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriAmichevoliCampoEsterno = Rec(3).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Sql = "SELECT 'MarcatoriCasaTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome " &
                    "Union All " &
                    "Select 'MarcatoriFuoriTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome " &
                    "Union All " &
                    "Select 'MarcatoriCampoEsternoTornei' As Cosa, Giocatori.Cognome, Giocatori.Nome, Count(*) As Goal, Giocatori.idGiocatore " &
                    "FROM(RisultatiAggiuntiviMarcatori Left Join Partite On Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                    "Left Join Giocatori On RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                    "Where RisultatiAggiuntiviMarcatori.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E' " &
                    "Group By Giocatori.Cognome, Giocatori.Nome"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "MarcatoriCasaTornei"
                                    NomiMarcatoriTorneiCasa.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriTorneiCasa = Rec(3).Value
                                Case "MarcatoriFuoriTornei"
                                    NomiMarcatoriTorneiFuori.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriTorneiFuori = Rec(3).Value
                                Case "MarcatoriCampoEsternoTornei"
                                    NomiMarcatoriTorneiCampoEsterno.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value)
                                    MarcatoriTorneiCampoEsterno = Rec(3).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Sql = "SELECT 'AvversariCasa' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'S' " &
                    "Union All " &
                    "SELECT 'AvversariFuori' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'N' " &
                    "Union All " &
                    "SELECT 'AvversariCampoEsterno' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteCampionatoIN & ") And Partite.Casa = 'E'"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "AvversariCasa"
                                    GoalAvvCampionatoCasa1Tempo = Rec(1).Value
                                    GoalAvvCampionatoCasa2Tempo = Rec(2).Value
                                    GoalAvvCampionatoCasa3Tempo = Rec(3).Value
                                Case "AvversariFuori"
                                    GoalAvvCampionatoFuori1Tempo = Rec(1).Value
                                    GoalAvvCampionatoFuori2Tempo = Rec(2).Value
                                    GoalAvvCampionatoFuori3Tempo = Rec(3).Value
                                Case "AvversariCampoEsterno"
                                    GoalAvvCampionatoCampoEsterno1Tempo = Rec(1).Value
                                    GoalAvvCampionatoCampoEsterno2Tempo = Rec(2).Value
                                    GoalAvvCampionatoCampoEsterno3Tempo = Rec(3).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Sql = "SELECT 'AvversariCasa' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'S' " &
                    "Union All " &
                    "SELECT 'AvversariFuori' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'N' " &
                    "Union All " &
                    "SELECT 'AvversariCampoEsterno' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteAmichevoliIN & ") And Partite.Casa = 'E'"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "AvversariCasa"
                                    GoalAvvAmichevoliCasa1Tempo = Rec(1).Value
                                    GoalAvvAmichevoliCasa2Tempo = Rec(2).Value
                                    GoalAvvAmichevoliCasa3Tempo = Rec(3).Value
                                Case "AvversariFuori"
                                    GoalAvvAmichevoliFuori1Tempo = Rec(1).Value
                                    GoalAvvAmichevoliFuori2Tempo = Rec(2).Value
                                    GoalAvvAmichevoliFuori3Tempo = Rec(3).Value
                                Case "AvversariCampoEsterno"
                                    GoalAvvAmichevoliCampoEsterno1Tempo = Rec(1).Value
                                    GoalAvvAmichevoliCampoEsterno2Tempo = Rec(2).Value
                                    GoalAvvAmichevoliCampoEsterno3Tempo = Rec(3).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Sql = "SELECT 'AvversariCasa' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'S' " &
                    "Union All " &
                    "SELECT 'AvversariFuori' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'N' " &
                    "Union All " &
                    "SELECT 'AvversariCampoEsterno' As Cosa, Sum(GoalAvvPrimoTempo) As PrimoTempo, Sum(GoalAvvSecondoTempo) As SecondoTempo, Sum(GoalAvvTerzoTempo) As TerzoTempo " &
                    "From RisultatiAggiuntivi Left Join Partite On RisultatiAggiuntivi.idPartita = Partite.idPartita " &
                    "Where RisultatiAggiuntivi.idPartita In (" & PartiteTorneiIN & ") And Partite.Casa = 'E'"
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Do Until Rec.Eof
                            Select Case Rec("Cosa").Value
                                Case "AvversariCasa"
                                    GoalAvvTorneiCasa1Tempo = Rec(1).Value
                                    GoalAvvTorneiCasa2Tempo = Rec(2).Value
                                    GoalAvvTorneiCasa3Tempo = Rec(3).Value
                                Case "AvversariFuori"
                                    GoalAvvTorneiFuori1Tempo = Rec(1).Value
                                    GoalAvvTorneiFuori2Tempo = Rec(2).Value
                                    GoalAvvTorneiFuori3Tempo = Rec(3).Value
                                Case "AvversariCampoEsterno"
                                    GoalAvvTorneiCampoEsterno1Tempo = Rec(1).Value
                                    GoalAvvTorneiCampoEsterno2Tempo = Rec(2).Value
                                    GoalAvvTorneiCampoEsterno3Tempo = Rec(3).Value
                            End Select

                            Rec.MoveNext
                        Loop
                        Rec.Close()
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
                            "Select Sum(Goal1Tempo) As G1Tempo, Sum(Goal2Tempo) As G2Tempo, Sum(Goal3Tempo) As G3Tempo, Sum(GA1Tempo) As GoalAvv1Tempo, Sum(Ga2Tempo) As GoalAvv2Tempo, Sum(GA3Tempo) As GoalAvv3Tempo, (Select RisGiochetti From RisultatiAggiuntivi Where idPartita =859)  As RisGiochetti From (" &
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
                                    If Rec("RisultatoATempi").Value = "N" Then
                                        Dim GoalTotaliFatti As Integer = Rec("G1Tempo").Value + Rec("G2Tempo").Value + Rec("G3Tempo").Value
                                        Dim GoalTotaliSubiti As Integer = Rec("GoalAvv1Tempo").Value + Rec("GoalAvv2Tempo").Value + Rec("GoalAvv3Tempo").Value
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
            End If
        End If

        If Ritorno = "" Then Ritorno = StringaErrore & " Nessun dato rilevato"

        Return Ritorno
    End Function

End Class