Public Class Encryption

    Private Shared Key As String = "AE6F39049135C519"

    Public Shared Function Encrypt(ByVal data As String) As String
        Try
            If data.Trim = "" Then
                Return ""
            End If

            Dim p As EncryptionClassLibrary.Encryption.Symmetric.Provider = EncryptionClassLibrary.Encryption.Symmetric.Provider.TripleDES

            Dim sym As New EncryptionClassLibrary.Encryption.Symmetric(p)
            sym.Key.Text = Key

            Dim encryptedData As EncryptionClassLibrary.Encryption.Data
            encryptedData = sym.Encrypt(New EncryptionClassLibrary.Encryption.Data(data))

            Return encryptedData.Base64
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function Decrypt(ByVal data As String) As String
        Try
            If data.Trim = "" Then
                Return ""
            End If

            Dim p As EncryptionClassLibrary.Encryption.Symmetric.Provider = EncryptionClassLibrary.Encryption.Symmetric.Provider.TripleDES

            Dim sym As New EncryptionClassLibrary.Encryption.Symmetric(p)
            sym.Key.Text = Key

            Dim cipherTextBytes As Byte() = Convert.FromBase64String(data)
            Dim ms As New System.IO.MemoryStream(cipherTextBytes)

            Dim decryptedData As EncryptionClassLibrary.Encryption.Data
            decryptedData = sym.Decrypt(ms)

            Return decryptedData.Text
        Catch ex As Exception
            Return ""
        End Try
    End Function

End Class
