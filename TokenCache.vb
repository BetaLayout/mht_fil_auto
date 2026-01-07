

Imports System.IO
Imports System.Security.Cryptography

Public Class TokenCache
    Private Shared ReadOnly _algorithm As New AesCryptoServiceProvider With {
       .KeySize = 256,
       .BlockSize = 128,
       .Mode = CipherMode.CBC,
       .Padding = PaddingMode.PKCS7
   }

    Private Shared _key As Byte() = Nothing
    Private Shared keyString As String = "O5mMjjrxWjCwMJLfIBVGgd81L8vNsfl+H+dMEe1PE/U="
    Public Shared Sub GenerateKey()
        Using aesAlg As New AesCryptoServiceProvider()
            aesAlg.KeySize = 256
            aesAlg.GenerateKey()
            _key = aesAlg.Key

            'keyString = Convert.ToBase64String(_key)  
            '//Need to put debul here and make keyString value above = to this
            'Dim berakpoint As String = "my_goodness"
        End Using
    End Sub

    Public Shared Function GetKey() As Byte()
        If keyString Is Nothing Then
            Throw New InvalidOperationException("Key has not been generated. Call GenerateKey() first.")
        End If
        Return Convert.FromBase64String(keyString)
    End Function

    Public Shared Sub EncryptToken(token As String)
        If keyString Is Nothing Then
            Throw New InvalidOperationException("Key has not been generated. Call GenerateKey() first.")
        End If

        Using aesAlg As New AesCryptoServiceProvider()
            aesAlg.Key = GetKey()

            ' Generate random IV
            aesAlg.GenerateIV()
            Dim iv = aesAlg.IV

            ' Combine IV and encrypted data
            Using ms As New MemoryStream()
                ms.Write(iv, 0, iv.Length)

                Using cs As New CryptoStream(ms,
                    aesAlg.CreateEncryptor(),
                    CryptoStreamMode.Write)
                    Using sw As New StreamWriter(cs)
                        sw.Write(token)
                    End Using
                End Using

                File.WriteAllBytes("S:\OMS_DATA\Applications\Access\encrypted_token.bin", ms.ToArray())
            End Using
        End Using
    End Sub

    Public Shared Function DecryptToken() As String
        If keyString Is Nothing Then
            Throw New InvalidOperationException("Key has not been generated. Call GenerateKey() first.")
        End If

        Dim combinedData = File.ReadAllBytes("S:\OMS_DATA\Applications\Access\encrypted_token.bin")

        Using aesAlg As New AesCryptoServiceProvider()
            ' Get IV from combined data
            Dim iv = New Byte(_algorithm.BlockSize \ 8 - 1) {}
            Array.Copy(combinedData, iv, iv.Length)
            aesAlg.IV = iv

            'Get encrypted data
            Dim encryptedData = New Byte(combinedData.Length - iv.Length - 1) {}
            Array.Copy(combinedData, iv.Length, encryptedData, 0, encryptedData.Length)
            aesAlg.Key = GetKey()

            Using ms As New MemoryStream(encryptedData)
                Using cs As New CryptoStream(ms,
                    aesAlg.CreateDecryptor(),
                    CryptoStreamMode.Read)
                    Using sr As New StreamReader(cs)
                        Return sr.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using
    End Function

End Class
