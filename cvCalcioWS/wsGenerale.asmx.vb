Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO

<System.Web.Services.WebService(Namespace:="http://cvcalcio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsGenerale
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function AggiornaDB(ByVal Numero As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida & ":" & Connessione
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Sql As String = ""

                Try
                    Select Case Numero
                        Case "1"
                            Sql = "Create Table CampiEsterni (idPartita Integer , Descrizione Text(255), CONSTRAINT TelefonatePK PRIMARY KEY (idPartita))"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Case "2"
                            Sql = "Create Table CoordinatePartite (idPartita Integer, Lat Text(15), Lon Text(15), CONSTRAINT CoordPK PRIMARY KEY (idPartita))"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Case "3"
                            Sql = "Create Table MeteoPartite (idPartita Integer, Tempo Text(30), Gradi Text(10), Umidita Text(10), Pressione Text(10), CONSTRAINT MeteoPK PRIMARY KEY (idPartita))"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Case "4"
                            Sql = "Create Table AvversariCoord (idAvversario Integer, Lat Text(30), Lon Text(30), CONSTRAINT AvvCoordPK PRIMARY KEY (idAvversario))"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Case "5"
                            Sql = "Alter Table CalendarioPartite Add Giocata Text(1)"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)

                            Sql = "Update CalendarioPartite Set Giocata='S'"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)

                            Sql = "Alter Table CalendarioDate Add idPartita Integer"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Case "6"
                            Sql = "Create Table Giornata (idUtente Integer, idAnno Integer, idCategoria Integer, idGiornata Integer, CONSTRAINT GiornataPK PRIMARY KEY (idUtente, idAnno, idCategoria))"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Case "7"
                            Sql = "Alter Table Anni Add NomeSquadra Text(50)"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)

                            Sql = "Create Table AnnoAttualeUtenti (idUtente Integer, idAnno Integer, CONSTRAINT AnnoAttualeUtentiPK PRIMARY KEY (idUtente))"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Case Else
                            Ritorno = StringaErrore & " Valore non valido"
                    End Select
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function RitornaVersioneApplicazione() As String
        Dim Ritorno As String = ""

        Dim gf As New GestioneFilesDirectory
        gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\NuoveVersioni\")
        Dim NuovaVersione As String = gf.LeggeFileIntero(Server.MapPath(".") & "\NuoveVersioni\Versione.txt")
        If NuovaVersione <> "" Then
            Ritorno = NuovaVersione
        Else
            Ritorno = StringaErrore & " Nessuna nuova versione rilevata"
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function RitornaAnni() As String
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

                Try
                    Sql = "SELECT * FROM Anni"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Ritorno = ""
                        Do Until Rec.Eof
                            Ritorno &= Rec("idAnno").Value & ";" &
                                Rec("Descrizione").Value & ";" &
                                Rec("NomeSquadra").Value & ";" &
                                "§"
                            Rec.MoveNext()
                        Loop
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
    Public Function RitornaValoriPerRegistrazione() As String
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
                Dim Anno As String

                If Now.Month >= 8 Then
                    ' Ci si sta registrando per l'anno in corso
                    Anno = Now.Year.ToString.Trim
                Else
                    ' Ci si sta registrando per l'anno in corso che è cominciato quello passato
                    Anno = (Now.Year - 1).ToString.Trim()
                End If

                Try
                    Sql = "SELECT * FROM Anni Where Descrizione Like '%" & Anno & "/%'"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        Ritorno = ""
                        Do Until Rec.Eof
                            Ritorno &= Rec("idAnno").Value & ";" &
                                Rec("Descrizione").Value & ";" &
                                Rec("NomeSquadra").Value & ";" &
                                "§"
                            Rec.MoveNext()
                        Loop
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
    Public Function RitornaTipologie() As String
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

                Try
                    Sql = "SELECT * FROM TipologiePartite Order By Descrizione"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " No tipologies found"
                        Else
                            Ritorno = ""
                            Do Until Rec.Eof
                                Ritorno &= Rec("idTipologia").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

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
    Public Function RitornaRuoli() As String
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

                Try
                    Sql = "SELECT * From Ruoli Order By idRuolo"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " No roles found"
                        Else
                            Ritorno = ""
                            Do Until Rec.Eof
                                Ritorno &= Rec("idRuolo").Value.ToString & ";" & Rec("Descrizione").Value.ToString & "§"

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
    Public Function RitornaMaxAnno() As String
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
                Dim idAnno As Integer

                Try
                    Sql = "SELECT Max(idAnno) From Anni"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " No years found"
                            idAnno = -1
                        Else
                            idAnno = Rec(0).Value
                            Ritorno = (Rec(0).Value) + 1 & ";"
                        End If
                        Rec.Close()
                    End If

                    If idAnno > -1 Then
                        Sql = "SELECT Descrizione From Anni Where idAnno=" & idanno
                        Rec = LeggeQuery(Conn, Sql, Connessione)
                        If TypeOf (Rec) Is String Then
                            Ritorno = Rec
                        Else
                            If Rec.Eof Then
                                Ritorno = StringaErrore & " No description found"
                            Else
                                Dim desc As String = Rec(0).Value
                                If desc.Contains("/") Then
                                    Dim c() As String = desc.Split("/")
                                    desc = Val(c(0) + 1).ToString & "/" & Val(c(1) + 1).ToString
                                End If

                                Ritorno &= desc & ";"
                            End If
                            Rec.Close()
                        End If
                    Else
                        Ritorno = StringaErrore & " Nessun anno rilevato"
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
    Public Function CreaNuovoAnno(idAnno As String, descAnno As String, nomeSquadra As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Sql As String = ""

                Try
                    Sql = "Insert Into Anni Values (" & idAnno & ", '" & descAnno.Replace(";", "_").Replace("'", "''") & "', '" & nomeSquadra.Replace(";", "_").Replace("'", "''") & "')"
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function RitornaAnnoAttualeUtente(idUtente As String) As String
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
                Dim idAnno As Integer

                Try
                    Sql = "SELECT * From AnnoAttualeUtenti Where idUtente=" & idUtente
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Rec.Close

                            Sql = "SELECT Max(idAnno) From Anni"
                            Rec = LeggeQuery(Conn, Sql, Connessione)
                            If TypeOf (Rec) Is String Then
                                Ritorno = Rec
                            Else
                                If Rec.Eof Then
                                    idAnno = 1
                                Else
                                    idAnno = Rec(0).Value
                                    Ritorno = (Rec(0).Value) & ";"
                                End If
                            End If
                        Else
                            idAnno = Rec(1).Value
                            Ritorno = (Rec(1).Value) & ";"
                        End If
                        Rec.Close()
                    End If

                    If idAnno > -1 Then
                        Sql = "SELECT Descrizione, NomeSquadra, Lat, Lon From Anni Where idAnno=" & idAnno
                        Rec = LeggeQuery(Conn, Sql, Connessione)
                        If TypeOf (Rec) Is String Then
                            Ritorno = Rec
                        Else
                            If Rec.Eof Then
                                Ritorno = StringaErrore & " No description found"
                            Else
                                Dim desc As String = Rec(0).Value
                                Dim NomeSquadra As String = "" & Rec("NomeSquadra").Value

                                Ritorno &= desc & ";" & NomeSquadra & ";" & Rec("Lat").Value & ";" & Rec("Lon").Value & ";"
                            End If
                            Rec.Close()
                        End If
                    Else
                        Ritorno = StringaErrore & " Nessun anno rilevato"
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
    Public Function ImpostaAnnoAttualeUtente(idAnno As String, idUtente As String) As String
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

                Try
                    Sql = "Delete From AnnoAttualeUtenti Where idUtente=" & idUtente
                    Ritorno = EsegueSql(Conn, Sql, Connessione)

                    Sql = "Insert Into AnnoAttualeUtenti Values (" & idUtente & ", " & idAnno & ")"
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function CreaPartitaHTML(idAnno As String, idPartita As String) As String
        Dim Ritorno As String = ""

        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida & ":" & Connessione
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                CreaHtmlPartita(Conn, Connessione, idAnno, idPartita)
            End If
        End If

        Return Ritorno
    End Function

End Class