Imports SelectPdf

Public Class pdfGest
	Public Function ConverteHTMLInPDF(NomeHtml As String, pathSalvataggio As String, pathLog As String)
		Dim Ritorno As String = ""

		Dim gf As New GestioneFilesDirectory
		Try
			Dim fileHtml As String = gf.LeggeFileIntero(NomeHtml)

			If pathLog <> "" Then
				gf.ApreFileDiTestoPerScrittura(pathLog)
				gf.ScriveTestoSuFileAperto("Conversione file " & NomeHtml)
				gf.ScriveTestoSuFileAperto("Salvataggio su " & pathSalvataggio)
				gf.ScriveTestoSuFileAperto("Contenuto " & fileHtml)
			End If
			gf.EliminaFileFisico(pathSalvataggio)
			'Dim pdf As PdfDocument = PdfGenerator.GeneratePdf(fileHtml, PageSize.A4)
			'pdf.Save(pathSalvataggio)
			SurroundingSub(fileHtml, pathSalvataggio)
			If pathLog <> "" Then
				gf.ScriveTestoSuFileAperto("Elaborazione effettuata")
			End If

			Ritorno = "*"
		Catch ex As Exception
			If pathLog <> "" Then
				gf.ScriveTestoSuFileAperto("Errore: " & ex.Message)
			End If
			Ritorno = StringaErrore & " " & ex.Message
		End Try
		If pathLog <> "" Then
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Return Ritorno
	End Function

	Private Sub SurroundingSub(htmlString As String, fileSalvataggio As String)
		' https://selectpdf.com/html-to-pdf/docs/html/PdfPageProperties.htm

		Dim converter As HtmlToPdf = New HtmlToPdf
		converter.Options.PdfPageSize = PdfPageSize.A4
		converter.Options.MarginLeft = 10
		converter.Options.MarginRight = 10
		converter.Options.MarginTop = 20
		converter.Options.MarginBottom = 20

		Dim doc As SelectPdf.PdfDocument = converter.ConvertHtmlString(htmlString)
		doc.Save(fileSalvataggio)
		doc.Close()
	End Sub
End Class
