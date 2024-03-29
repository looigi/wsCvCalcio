﻿Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://templates.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsTemplates
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function eliminaFileScheletroAssociato(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\associato.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroEMailAss(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\mail_associato.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroSollecito(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\mail_sollecito.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroMail(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_mail.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroIscrizione(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_iscrizione_" & Squadra & ".txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroConvocazione(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\nuova_partita.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroPrivacy(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_privacy.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroFirma(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_firma.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroRicevutaStandard(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\ricevuta_pagamento.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroRicevutaScontrino(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\ricevuta_scontrino.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function eliminaFileScheletroTestoAggConv(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\testo_convocazioni.txt"
		Dim Ritorno As String = "*"
		If ControllaEsistenzaFile(pathFilePosta) Then
			File.Delete(pathFilePosta)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroAssociato(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\associato.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\associato.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroTestoAggConv(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\testo_convocazioni.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\testo_convocazioni.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroEMailAss(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\mail_associato.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\mail_associato.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroRicevutaScontrino(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\ricevuta_scontrino.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\ricevuta_scontrino.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroRicevutaStandard(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\ricevuta_pagamento.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\ricevuta_pagamento.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroFirma(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_firma.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_firma.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroConvocazione(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\nuova_partita.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\nuova_partita.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroSollecito(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\mail_sollecito.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\mail_sollecito.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroMail(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_mail.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_mail.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroIscrizione(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_iscrizione_" & Squadra & ".txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_iscrizione_" & Squadra & ".txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomeFileScheletroPrivacy(Squadra As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim filePaths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim p() As String = filePaths.Split(";")
		If Strings.Right(p(0), 1) <> "\" Then
			p(0) &= "\"
		End If
		Dim pathFilePosta As String = p(0) & Squadra & "\Scheletri\base_privacy.txt"
		Dim Ritorno As String = "MODIFICATO"
		If Not ControllaEsistenzaFile(pathFilePosta) Then
			pathFilePosta = HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_privacy.txt"
			Ritorno = "ORIGINALE"
		End If

		Return Ritorno
	End Function

End Class