Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_ute.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsUtenti
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function RitornaUtentePerLogin(ByVal idAnno As String, Utente As String, Password As String) As String
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
                    ' Sql = "SELECT * FROM UtentiMobile Where idAnno=" & idAnno & " And Utente='" & Utente.Replace("'", "''") & "'"
                    Sql = "SELECT UtentiMobile.idAnno, idUtente, Utente, Cognome, Nome, " &
                        "Password, EMail, Categorie.idCategoria As idCat1, idTipologia, Categorie.Descrizione As Descr1 " &
                        "FROM (UtentiMobile " &
                        "Left Join Categorie On UtentiMobile.idCategoria=Categorie.idCategoria And UtentiMobile.idAnno=Categorie.idAnno) " &
                        "Where Utente='" & Utente.Replace("'", "''") & "'"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun utente rilevato"
                        Else
                            If Password <> Rec("Password").Value.ToString Then
                                Ritorno = StringaErrore & " Password non valida"
                            Else
                                Ritorno = ""
                                Do Until Rec.Eof
                                    Ritorno &= Rec("idAnno").Value & ";" &
                                        Rec("idUtente").Value & ";" &
                                        Rec("Utente").Value & ";" &
                                        Rec("Cognome").Value & ";" &
                                        Rec("Nome").Value & ";" &
                                        Rec("Password").Value & ";" &
                                        Rec("EMail").Value & ";" &
                                        Rec("idCat1").Value & ";" &
                                        Rec("idTipologia").Value & ";" &
                                        Rec("Descr1").Value & ";" &
                                        "§"
                                    Rec.MoveNext()
                                Loop
                            End If
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
    Public Function RitornaUtenteDaID(ByVal idAnno As String, idUtente As String) As String
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
                    ' Sql = "SELECT * FROM UtentiMobile Where idAnno=" & idAnno & " And idUtente=" & idUtente
                    Sql = "SELECT UtentiMobile.*, Categorie.Descrizione " &
                        "From UtentiMobile LEFT Join Categorie On (UtentiMobile.idCategoria = Categorie.idCategoria) And (UtentiMobile.idAnno = Categorie.idAnno) " &
                        "Where idUtente = " & idUtente
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun utente rilevato"
                        Else
                            Ritorno = ""
                            Do Until Rec.Eof
                                Ritorno &= Rec("idAnno").Value & ";" &
                                    Rec("idUtente").Value & ";" &
                                    Rec("Utente").Value & ";" &
                                    Rec("Cognome").Value & ";" &
                                    Rec("Nome").Value & ";" &
                                    Rec("Password").Value & ";" &
                                    Rec("EMail").Value & ";" &
                                    Rec("idCategoria").Value & ";" &
                                    Rec("idTipologia").Value & ";" &
                                    Rec("Descrizione").Value & ";" &
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
    Public Function RitornaListaUtenti(idAnno As String) As String
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
                    Sql = "SELECT UtentiMobile.idAnno, UtentiMobile.idUtente, UtentiMobile.Utente, UtentiMobile.Cognome, UtentiMobile.Nome, UtentiMobile.EMail, Categorie.Descrizione As Categoria, " &
                        "UtentiMobile.idTipologia, UtentiMobile.Password, Categorie.idCategoria " &
                        "FROM (UtentiMobile LEFT JOIN Categorie ON UtentiMobile.idCategoria = Categorie.idCategoria And UtentiMobile.idAnno = Categorie.idAnno) " &
                        "Order By 2,1;"
                    ' "Where UtentiMobile.idAnno=" & idAnno & " Order By 2,1;"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun utente rilevato"
                        Else
                            Ritorno = ""
                            Do Until Rec.Eof
                                Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")

                                Sql = " Select * From AnnoAttualeUtenti Where idUtente=" & Rec("idUtente").Value
                                Rec2 = LeggeQuery(Conn, Sql, Connessione)
                                Dim AnnoUtente As Integer = Rec2("idAnno").Value
                                Rec2.Close

                                Sql = " Select * From Anni Where idAnno=" & AnnoUtente.ToString
                                Rec2 = LeggeQuery(Conn, Sql, Connessione)
                                Dim NomeSquadra As String = Rec2("NomeSquadra").Value
                                Rec2.Close

                                Ritorno &= "0;" &
                                    Rec("idUtente").Value & ";" &
                                    Rec("Utente").Value & ";" &
                                    Rec("Cognome").Value & ";" &
                                    Rec("Nome").Value & ";" &
                                    Rec("EMail").Value & ";" &
                                    NomeSquadra & ";" &
                                    Rec("idTipologia").Value & ";" &
                                    Rec("Password").Value & ";" &
                                    Rec("idCategoria").Value & ";" &
                                    Rec("Categoria").Value & ";" &
                                    "§"

                                Rec.MoveNext()
                            Loop
                        End If
                        Rec.Close()

                        'Ritorno &= "£"

                        'Sql = "Select * From Categorie Where idAnno=" & idAnno & " And Eliminato = 'N' Order By Ordinamento"
                        'Rec = LeggeQuery(Conn, Sql, Connessione)
                'If TypeOf (Rec) Is String Then
                '    Ritorno = Rec
                'Else
                '    If Rec.Eof Then
                '        Ritorno = StringaErrore & " Nessuna categoria rilevata"
                '    Else
                '        Do Until Rec.Eof
                '            Ritorno &= Rec("idCategoria").Value & ";" &
                '                Rec("Descrizione").Value & ";" &
                '                "§"
                '            Rec.MoveNext()
                '        Loop
                '    End If
                '    Rec.Close()
                'End If

                'Ritorno &= "£"
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
    Public Function SalvaUtente(ByVal idAnno As String, Utente As String, Cognome As String, Nome As String, EMail As String,
                                Password As String, idCategoria As String, idTipologia As String) As String
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
                Dim idUtente As String = ""

				Try
					Sql = "SELECT * FROM UtentiMobile Where Ucase(Trim(Utente))='" & Utente.Trim.ToUpper & "' And idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Sql = "SELECT Max(idUtente)+1 FROM UtentiMobile" ' Where idAnno=" & idAnno
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = StringaErrore & " Nessun utente rilevato"
								Else
									If Rec(0).Value Is DBNull.Value Then
										idUtente = "1"
									Else
										idUtente = Rec(0).Value.ToString
									End If
								End If
								Rec.Close()
							End If

							If idUtente <> "" Then
								Sql = "Insert Into UtentiMobile Values (" &
							" " & idAnno & ", " &
							" " & idUtente & ", " &
							"'" & Utente & "', " &
							"'" & Cognome & "', " &
							"'" & Nome & "', " &
							"'" & Password & "', " &
							"'" & EMail & "', " &
							" " & idCategoria & ", " &
							" " & idTipologia & " " &
							")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)

								Try
									Sql = "Delete From AnnoAttualeUtenti Where idUtente=" & idUtente
									Ritorno = EsegueSql(Conn, Sql, Connessione)

									Sql = "Insert Into AnnoAttualeUtenti Values (" & idUtente & ", " & idAnno & ")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
								Catch ex As Exception
									Ritorno = StringaErrore & " " & ex.Message
								End Try

								If Ritorno = "*" Then Ritorno = idUtente
							Else
								Ritorno = StringaErrore & " Problemi nel rilevamento dell'ID Utente"
							End If
						Else
							Ritorno = StringaErrore & " Utente già esistente per l'anno in corso"
						End If
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
    Public Function ModificaUtente(ByVal idAnno As String, Utente As String, Cognome As String, Nome As String, EMail As String,
                                Password As String, idCategoria As String, idTipologia As String, idUtente As String) As String
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

                ' Sql = "Delete From UtentiMobile Where idAnno=" & idAnno & " And idUtente=" & idUtente
                Sql = "Delete From UtentiMobile Where idUtente=" & idUtente
                Ritorno = EsegueSql(Conn, Sql, Connessione)

                Try
                    Sql = "Insert Into UtentiMobile Values (" &
                        "" & idAnno & ", " &
                        "" & idUtente & ", " &
                        "'" & Utente & "', " &
                        "'" & Cognome & "', " &
                        "'" & Nome & "', " &
                        "'" & Password & "', " &
                        "'" & EMail & "', " &
                        " " & idCategoria & ", " &
                        "" & idTipologia & ")"
                    Ritorno = EsegueSql(Conn, Sql, Connessione)

                    Try
                        Sql = "Delete From AnnoAttualeUtenti Where idUtente=" & idUtente
                        Ritorno = EsegueSql(Conn, Sql, Connessione)

                        Sql = "Insert Into AnnoAttualeUtenti Values (" & idUtente & ", " & idAnno & ")"
                        Ritorno = EsegueSql(Conn, Sql, Connessione)
                    Catch ex As Exception
                        Ritorno = StringaErrore & " " & ex.Message
                    End Try

                    If Ritorno = "*" Then Ritorno = idUtente
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function
End Class