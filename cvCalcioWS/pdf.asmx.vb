Imports System.Web.Services
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://pdf.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class pdfClass
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function ConverteHTMLToPDF(NomeHtml As String, PathSalvataggio As String) As String
		Dim p As New pdfGest

		Return p.ConverteHTMLInPDF(Server.MapPath("."), NomeHtml, PathSalvataggio, "")
	End Function

End Class