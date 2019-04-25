﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_all.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsAllenatori
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function SalvaAllenatore(idAnno As String, idCategoria As String, idAllenatore As String, Cognome As String, Nome As String, EMail As String, Telefono As String) As String
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
                Dim idAll As Integer = -1

                If idAllenatore = "-1" Then
                    Try
                        Sql = "SELECT Max(idAllenatore)+1 FROM Allenatori Where idAnno=" & idAnno
                        Rec = LeggeQuery(Conn, Sql, Connessione)
                        If TypeOf (Rec) Is String Then
                            Ritorno = Rec
                        Else
                            If Rec(0).Value Is DBNull.Value Then
                                idAll = 1
                            Else
                                idAll = Rec(0).Value
                            End If
                            Rec.Close()
                        End If
                    Catch ex As Exception
                        Ritorno = StringaErrore & " " & ex.Message
                    End Try
                Else
                    idAll = idAllenatore
                    Sql = "Delete * From Allenatori Where idAnno=" & idAnno & " And idAllenatore=" & idAll
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                End If

                Sql = "Insert Into Allenatori Values (" &
                    " " & idAll & ", " &
                    "'" & Cognome.Replace("'", "''") & "', " &
                    "'" & Nome.Replace("'", "''") & "', " &
                    "'" & EMail.Replace("'", "''") & "', " &
                    "'" & Telefono.Replace("'", "''") & "', " &
                    "'N', " &
                    " " & idCategoria & ", " &
                    " " & idAnno & " " &
                    ")"
                Ritorno = EsegueSql(Conn, Sql, Connessione)

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function RitornaAllenatoriCategoria(ByVal idAnno As String, idCategoria As String) As String
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
                    Sql = "SELECT * FROM Allenatori Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Eliminato='N' Order By Cognome, Nome"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun allenatore rilevato"
                        Else
                            Ritorno = ""
                            Do Until Rec.Eof
                                Ritorno &= Rec("idAllenatore").Value & ";" &
                                    Rec("Cognome").Value.ToString.Trim & ";" &
                                    Rec("Nome").Value.ToString.Trim & ";" &
                                    Rec("EMail").Value.ToString.Trim & ";" &
                                    Rec("Telefono").Value.ToString.Trim & ";" &
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
    Public Function EliminaAllenatore(ByVal idAnno As String, idAllenatore As String) As String
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
                Dim Ok As Boolean = True

                Try
                    Sql = "Update Allenatori Set Eliminato='S' Where idAnno=" & idAnno & " And idAllenatore=" & idAllenatore
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                    Ok = False
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function
End Class