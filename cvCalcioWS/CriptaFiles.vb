Imports System.IO
Imports System.Security.Cryptography

Public Class CriptaFiles
	' Password standard = WPippoBaudo227!

	Public Sub EncryptFile(ByVal password As String, ByVal _
		in_file As String, ByVal out_file As String)
		CryptFile(password, in_file, out_file, True)
	End Sub

	Public Sub DecryptFile(ByVal password As String, ByVal _
		in_file As String, ByVal out_file As String)
		CryptFile(password, in_file, out_file, False)
	End Sub

	Public Sub CryptFile(ByVal password As String, ByVal _
		in_file As String, ByVal out_file As String, ByVal _
		encrypt As Boolean)
		' Create input and output file streams.
		Using in_stream As New FileStream(in_file,
			FileMode.Open, FileAccess.Read)
			Using out_stream As New FileStream(out_file,
				FileMode.Create, FileAccess.Write)
				' Encrypt/decrypt the input stream into the
				' output stream.
				CryptStream(password, in_stream, out_stream,
					encrypt)
			End Using
		End Using
	End Sub

	Private Sub MakeKeyAndIV(ByVal password As String, ByVal _
		salt() As Byte, ByVal key_size_bits As Integer, ByVal _
		block_size_bits As Integer, ByRef key() As Byte, ByRef _
		iv() As Byte)
		Dim derive_bytes As New Rfc2898DeriveBytes(password,
		salt, 1000)

		key = derive_bytes.GetBytes(key_size_bits / 8)
		iv = derive_bytes.GetBytes(block_size_bits / 8)
	End Sub

	Private Sub CryptStream(ByVal password As String, ByVal _
		in_stream As Stream, ByVal out_stream As Stream, ByVal _
		encrypt As Boolean)
		' Make an AES service provider.
		Dim aes_provider As New AesCryptoServiceProvider()

		' Find a valid key size for this provider.
		Dim key_size_bits As Integer = 0
		For i As Integer = 1024 To 1 Step -1
			If (aes_provider.ValidKeySize(i)) Then
				key_size_bits = i
				Exit For
			End If
		Next i
		Debug.Assert(key_size_bits > 0)
		'Console.WriteLine("Key size: " & key_size_bits)

		' Get the block size for this provider.
		Dim block_size_bits As Integer = aes_provider.BlockSize

		' Generate the key and initialization vector.
		Dim key() As Byte = Nothing
		Dim iv() As Byte = Nothing
		Dim salt() As Byte = {&H0, &H0, &H1, &H2, &H3, &H4,
			&H5, &H6, &HF1, &HF0, &HEE, &H21, &H22, &H45}
		MakeKeyAndIV(password, salt, key_size_bits,
			block_size_bits, key, iv)

		' Make the encryptor or decryptor.
		Dim crypto_transform As ICryptoTransform
		If (encrypt) Then
			crypto_transform =
				aes_provider.CreateEncryptor(key, iv)
		Else
			crypto_transform =
				aes_provider.CreateDecryptor(key, iv)
		End If

		' Attach a crypto stream to the output stream.
		' Closing crypto_stream sometimes throws an
		' exception if the decryption didn't work
		' (e.g. if we use the wrong password).
		Try
			Using crypto_stream As New CryptoStream(out_stream,
				crypto_transform, CryptoStreamMode.Write)
				' Encrypt or decrypt the file.
				Const block_size As Integer = 1024
				Dim buffer(block_size) As Byte
				Dim bytes_read As Integer
				Do
					' Read some bytes.
					bytes_read = in_stream.Read(buffer, 0,
						block_size)
					If (bytes_read = 0) Then Exit Do

					' Write the bytes into the CryptoStream.
					crypto_stream.Write(buffer, 0, bytes_read)
				Loop
			End Using
		Catch
		End Try

		crypto_transform.Dispose()
	End Sub
End Class
