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

End Class