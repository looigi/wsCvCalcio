Imports System.Windows.Forms
Imports SelectPdf

Public Class pdfGest
	Public Function ConverteHTMLInPDF(NomeHtml As String, pathSalvataggio As String, pathLog As String, Optional noMargini As Boolean = False, Optional Orizzontale As Boolean = False, Optional AltezzaReport As Integer = -1)
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
			SurroundingSub(fileHtml, pathSalvataggio, noMargini, Orizzontale, AltezzaReport)
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

	Private Sub SurroundingSub(htmlString As String, fileSalvataggio As String, noMargini As Boolean, Orizzontale As Boolean, AltezzaReport As Integer)
		' https://selectpdf.com/html-to-pdf/docs/html/PdfPageProperties.htm

		Dim converter As HtmlToPdf = New HtmlToPdf
		' noMargini = False
		If noMargini = False Then
			If AltezzaReport > -1 Then
				converter.Options.PdfPageSize = PdfPageSize.Custom
				converter.Options.PdfPageCustomSize = New Drawing.SizeF(210, AltezzaReport)
			Else
				converter.Options.PdfPageSize = PdfPageSize.A4
				If Orizzontale = True Then
					converter.Options.PdfPageOrientation = PdfPageOrientation.Landscape
				End If
			End If

			converter.Options.JpegCompressionEnabled = True
			converter.Options.MarginLeft = 50
			converter.Options.MarginRight = 50
			converter.Options.MarginTop = 50
			converter.Options.MarginBottom = 50

			converter.Footer.Height = 30
			converter.Options.DisplayFooter = True
			converter.Footer.DisplayOnFirstPage = True
			converter.Footer.DisplayOnOddPages = True
			converter.Footer.DisplayOnEvenPages = True

			Dim Datella As String = Format(Now.Day, "00") & "-" & Format(Now.Month, "00") & "-" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
			Dim textData As New PdfTextSection(0, 10, datella, New System.Drawing.Font("Arial", 8))
			textData.HorizontalAlign = PdfTextHorizontalAlign.Left
			Dim text As New PdfTextSection(0, 10, "Stampato tramite InCalcio – www.incalcio.it – info@incalcio.it", New System.Drawing.Font("Arial", 8))
			text.HorizontalAlign = PdfTextHorizontalAlign.Center
			Dim textPagina As New PdfTextSection(0, 10, "Pagina: {page_number} di {total_pages}  ", New System.Drawing.Font("Arial", 8))
			textPagina.HorizontalAlign = PdfTextHorizontalAlign.Right
			converter.Footer.Add(textData)
			converter.Footer.Add(text)
			converter.Footer.Add(textPagina)
		Else
			If AltezzaReport > -1 Then
				Dim alte As Single = AltezzaReport * 0.264583333

				converter.Options.PdfPageSize = PdfPageSize.Custom
				converter.Options.PdfPageCustomSize = New Drawing.SizeF(210, alte)
			End If

			converter.Options.JpegCompressionEnabled = True
			converter.Options.MarginLeft = 2
			converter.Options.MarginRight = 2
			converter.Options.MarginTop = 2
			converter.Options.MarginBottom = 2
		End If

		Dim doc As SelectPdf.PdfDocument = converter.ConvertHtmlString(htmlString)
		doc.Save(fileSalvataggio)
		doc.Close()
	End Sub
End Class
