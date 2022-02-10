Module mdlTrasfSMS
	Public ultimaNotifica As String = ""
	Public ultimoMessaggio As String = ""
	Public ultimaPosizione As String = ""

	Public Function SistemaStringa(ByVal Cosa As String) As String
		Dim Ritorno As String = Cosa

		Ritorno = MetteMaiuscoleDopoOgniSpazio(Ritorno)
		Ritorno = RaddoppiaApici(Ritorno)

		Return Ritorno
	End Function

	Public Function MetteMaiuscoleDopoOgniSpazio(ByVal Cosa As String) As String
		Dim Appoggio As String
		Dim Ritorno As String
		Dim I As Integer

		Appoggio = LCase(Cosa)
		Ritorno = UCase(Mid(Appoggio, 1, 1))
		For I = 2 To Len(Appoggio)
			Ritorno = Ritorno & Mid(Appoggio, I, 1)
			If Mid(Appoggio, I, 1) = " " Or Mid(Appoggio, I, 1) = "'" Then
				Ritorno = Ritorno & UCase(Mid(Appoggio, I + 1, 1))
				I = I + 1
			End If
		Next I

		MetteMaiuscoleDopoOgniSpazio = Ritorno
	End Function

	Public Function RaddoppiaApici(ByVal Cosa As String) As String
		Dim Ritorno As String = Cosa.Trim

		Ritorno = Ritorno.Replace("'", "''")

		Return Ritorno
	End Function

End Module

