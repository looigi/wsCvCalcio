Public Class CodiceFiscale
    Public Function CreaCodiceFiscale(tCognome As String, tNome As String,
                                      tDataNascita As String, tLuogoNascita As String, Maschio As Boolean) As String
        Dim StringaMese As String = "ABCDEHLMPRST"
        Dim Cognome As String = tCognome
        Dim Nome As String = tNome
        Dim CognNome As String = PrendeLettereCF(Cognome.ToUpper, Nome.ToUpper)
        Dim sDataNascita As String = tDataNascita
        Dim DataNascita As Date
        Dim Giorno As String = ""
        Dim Mese As String = ""
        Dim Anno As String = ""
        If sDataNascita <> "" Then
            If IsDate(sDataNascita) Then
                DataNascita = sDataNascita
                If Maschio = True Then
                    Giorno = DataNascita.Day.ToString.Trim
                Else
                    Giorno = (DataNascita.Day + 40).ToString.Trim
                End If
                Mese = Mid(StringaMese, DataNascita.Month, 1)
                Anno = DataNascita.Year.ToString.Trim
                If Giorno.Length = 1 Then
                    Giorno = "0" & Giorno
                End If
                Anno = Mid(Anno, 3, 2)
            End If
        End If
        Dim Comune As String = tLuogoNascita
        Dim CodComune As String = ""
        If Comune <> "" Then
            CodComune = PrendeCodiceComuneperCF(Comune)
        End If

        Dim Cf As String = CognNome & Anno & Mese & Giorno & CodComune

        Dim CodControllo As String = ""
        If Cf.Length = 15 Then
            Dim Valore As Integer = 0
            Dim ValoreTotale As Integer = 0
            Dim Carattere As String = ""

            For i As Integer = 2 To 14 Step 2
                Carattere = Mid(Cf, i, 1).ToUpper
                Select Case Carattere
                    Case "A", "0"
                        Valore = 0
                    Case "B", "1"
                        Valore = 1
                    Case "C", "2"
                        Valore = 2
                    Case "D", "3"
                        Valore = 3
                    Case "E", "4"
                        Valore = 4
                    Case "F", "5"
                        Valore = 5
                    Case "G", "6"
                        Valore = 6
                    Case "H", "7"
                        Valore = 7
                    Case "I", "8"
                        Valore = 8
                    Case "J", "9"
                        Valore = 9
                    Case "K"
                        Valore = 10
                    Case "L"
                        Valore = 11
                    Case "M"
                        Valore = 12
                    Case "N"
                        Valore = 13
                    Case "O"
                        Valore = 14
                    Case "P"
                        Valore = 15
                    Case "Q"
                        Valore = 16
                    Case "R"
                        Valore = 17
                    Case "S"
                        Valore = 18
                    Case "T"
                        Valore = 19
                    Case "U"
                        Valore = 20
                    Case "V"
                        Valore = 21
                    Case "W"
                        Valore = 22
                    Case "X"
                        Valore = 23
                    Case "Y"
                        Valore = 24
                    Case "Z"
                        Valore = 25
                    Case Else
                        Stop
                End Select
                ValoreTotale += Valore
            Next
            For i As Integer = 1 To 15 Step 2
                Carattere = Mid(Cf, i, 1).ToUpper
                Select Case Carattere
                    Case "A", "0"
                        Valore = 1
                    Case "B", "1"
                        Valore = 0
                    Case "C", "2"
                        Valore = 5
                    Case "D", "3"
                        Valore = 7
                    Case "E", "4"
                        Valore = 9
                    Case "F", "5"
                        Valore = 13
                    Case "G", "6"
                        Valore = 15
                    Case "H", "7"
                        Valore = 17
                    Case "I", "8"
                        Valore = 19
                    Case "J", "9"
                        Valore = 21
                    Case "K"
                        Valore = 2
                    Case "L"
                        Valore = 4
                    Case "M"
                        Valore = 18
                    Case "N"
                        Valore = 20
                    Case "O"
                        Valore = 11
                    Case "P"
                        Valore = 3
                    Case "Q"
                        Valore = 6
                    Case "R"
                        Valore = 8
                    Case "S"
                        Valore = 12
                    Case "T"
                        Valore = 14
                    Case "U"
                        Valore = 16
                    Case "V"
                        Valore = 10
                    Case "W"
                        Valore = 22
                    Case "X"
                        Valore = 25
                    Case "Y"
                        Valore = 24
                    Case "Z"
                        Valore = 23
                    Case Else
                        Stop
                End Select
                ValoreTotale += Valore
            Next
            ValoreTotale = ValoreTotale Mod 26
            Dim CarattereSpeciale As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            CodControllo = Mid(CarattereSpeciale, ValoreTotale + 1, 1)
        End If

        Return (Cf & CodControllo).ToUpper
    End Function

    Public Function PrendeCodiceComuneperCF(Quale As String) As String
        'Dim Ritorno As String = ""
        'Dim DB As New GestioneACCESS

        'If DB.LeggeImpostazioniDiBase("ConnDB") = True Then
        '    Dim ConnSQL As Object = DB.ApreDB()
        '    Dim Rec As Object = CreateObject("ADODB.Recordset")
        '    Dim Sql As String

        '    Sql = "Select * From tListaComuniPerCF Where LTrim(Rtrim(Upper(Comune)))='" & Quale.Replace("'", "''") & "'"
        '    Rec = DB.LeggeQuery(ConnSQL, Sql)
        '    If Rec.Eof = True Then
        '        Ritorno = ""
        '    Else
        '        Ritorno = Rec("CodFisco").Value
        '    End If
        '    Rec.Close()

        '    ConnSQL.close()
        '    ConnSQL = Nothing
        'End If

        'DB = Nothing

        'Return Ritorno

        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(HttpContext.Current.Server.MapPath("."), "")

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = CreateObject("ADODB.Recordset")
                Dim Sql As String = ""

                Sql = "Select * From ComuniItaliani Where Ltrim(Rtrim(Upper(Comune))) = '" & Quale.ToUpper.Trim & "'"
                Rec = LeggeQuery(Conn, Sql, Connessione)
                If TypeOf (Rec) Is String Then
                    Ritorno = Rec
                Else
                    If Rec.Eof Then
                        Ritorno = "-----"
                    Else
                        Ritorno = Rec("CodiceCatastale").Value
                    End If
                    Rec.Close()
                End If
            End If
        End If

        Return Ritorno
    End Function

    Private Function PrendeLettereCF(ByVal Cognome As String, ByVal Nome As String) As String
        Dim fCalcolaCodiceFiscale As String = ""

        Dim lCiclo1 As Byte
        Dim lCarattere As String
        Dim lConsonanti As String
        Dim lVocali As String

        ' *** 3 caratteri estratti dal COGNOME ***

        ' Separa le CONSONANTI dalle VOCALI
        ' SPAZI ed ACCENTI vengono scartati
        lConsonanti = ""
        lVocali = ""
        lCarattere = ""
        For lCiclo1 = 1 To Len(Cognome)
            lCarattere = Mid(Cognome, lCiclo1, 1)
            Select Case lCarattere
                Case "A", "E", "I", "O", "U"
                    lVocali = lVocali + lCarattere
                Case "B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "V", "W", "X", "Y", "Z"
                    lConsonanti = lConsonanti + lCarattere
            End Select
        Next

        If Len(lConsonanti) > 2 Then
            ' 3 o + consonanti
            fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                Left(lConsonanti, 3) ' COGNOME OK
        ElseIf Len(lConsonanti) = 2 Then
            ' 2 consonanti
            If Len(lVocali) > 0 Then
                ' 1 o + vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    Left(lConsonanti, 2) &
                    Left(lVocali, 1) ' COGNOME OK
            Else
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lConsonanti &
                    "X" ' COGNOME OK
            End If
        ElseIf Len(lConsonanti) = 1 Then
            ' 1 consonante
            If Len(lVocali) > 1 Then
                ' 2 o + vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lConsonanti &
                    Left(lVocali, 2) ' COGNOME OK
            ElseIf Len(lVocali) = 1 Then
                ' 1 vocale
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lConsonanti &
                    lVocali &
                    "X" ' COGNOME OK
            Else
                ' 0 vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lConsonanti &
                    "XX" ' COGNOME OK
            End If
        Else
            ' Nessuna consonante
            If Len(lVocali) > 2 Then
                ' 3 o + vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                Left(lVocali, 3) ' COGNOME OK
            ElseIf Len(lVocali) = 2 Then
                ' 2 vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lVocali &
                    "X" ' COGNOME OK
            ElseIf Len(lVocali) = 1 Then
                ' 1 vocale
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lVocali &
                    "XX" ' COGNOME OK
            Else
                ' 0 vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    "XXX" ' COGNOME OK
            End If
        End If

        ' *** 3 caratteri estratti dal NOME ***

        ' Separa le CONSONANTI dalle VOCALI
        ' SPAZI ed ACCENTI vengono scartati
        lConsonanti = ""
        lVocali = ""
        lCarattere = ""
        For lCiclo1 = 1 To Len(Nome)
            lCarattere = Mid(Nome, lCiclo1, 1)
            Select Case lCarattere
                Case "A", "E", "I", "O", "U"
                    lVocali = lVocali + lCarattere
                Case "B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "V", "W", "X", "Y", "Z"
                    lConsonanti = lConsonanti + lCarattere
            End Select
        Next

        If Len(lConsonanti) > 3 Then
            ' 4 o + consonanti
            fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                Left(lConsonanti, 1) &
                Mid(lConsonanti, 3, 2) ' NOME OK
        ElseIf Len(lConsonanti) = 3 Then
            ' 3 consonanti
            fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                Left(lConsonanti, 3) ' NOME OK
        ElseIf Len(lConsonanti) = 2 Then
            ' 2 consonanti
            If Len(lVocali) > 0 Then
                ' 1 o + vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    Left(lConsonanti, 2) &
                    Left(lVocali, 1) ' NOME OK
            Else
                ' 0 vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    Left(lConsonanti, 2) &
                    "X" ' NOME OK
            End If
        ElseIf Len(lConsonanti) = 1 Then
            ' 1 consonante
            If Len(lVocali) > 1 Then
                ' 2 o + vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lConsonanti &
                    Left(lVocali, 2) ' NOME OK
            ElseIf Len(lVocali) = 1 Then
                ' 1 vocale
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lConsonanti &
                    lVocali &
                    "X" ' NOME OK
            Else
                ' 0 vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lConsonanti &
                    "XX" ' NOME OK
            End If
        Else
            ' Nessuna consonante
            If Len(lVocali) > 2 Then
                ' 3 o + vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    Left(lVocali, 3) ' NOME OK
            ElseIf Len(lVocali) = 2 Then
                ' 2 vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    Left(lVocali, 2) &
                    "X" ' NOME OK
            ElseIf Len(lVocali) = 1 Then
                ' 1 vocale
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    lVocali &
                    "XX" ' NOME OK                
            Else
                ' 0 vocali
                fCalcolaCodiceFiscale = fCalcolaCodiceFiscale &
                    "XXX" ' NOME OK
            End If
        End If

        Return fCalcolaCodiceFiscale
    End Function

End Class
