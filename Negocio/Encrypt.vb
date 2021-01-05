Imports System.Text
Imports System.Security.Cryptography
Imports System.IO

Public Class Encrypt
    Public Shared Function Encrypt_QueryString(ByVal str As String) As String
        Dim EncrptKey As String = "2013;[pnuLIT)WebCodeExpert"
        Dim byKey As Byte() = {}
        Dim IV As Byte() = {18, 52, 86, 120, 144, 171, _
         205, 239}
        byKey = System.Text.Encoding.UTF8.GetBytes(EncrptKey.Substring(0, 8))
        Dim des As New DESCryptoServiceProvider()
        Dim inputByteArray As Byte() = Encoding.UTF8.GetBytes(str)
        Dim ms As New MemoryStream()
        Dim cs As New CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write)
        cs.Write(inputByteArray, 0, inputByteArray.Length)
        cs.FlushFinalBlock()
        Return Convert.ToBase64String(ms.ToArray())
    End Function
    Public Shared Function Decrypt_QueryString(ByVal str As String) As String
        str = str.Replace(" ", "+")
        Dim DecryptKey As String = "2013;[pnuLIT)WebCodeExpert"
        Dim byKey As Byte() = {}
        Dim IV As Byte() = {18, 52, 86, 120, 144, 171, _
         205, 239}
        Dim inputByteArray As Byte() = New Byte(str.Length - 1) {}

        byKey = System.Text.Encoding.UTF8.GetBytes(DecryptKey.Substring(0, 8))
        Dim des As New DESCryptoServiceProvider()
        inputByteArray = Convert.FromBase64String(str)
        Dim ms As New MemoryStream()
        Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write)
        cs.Write(inputByteArray, 0, inputByteArray.Length)
        cs.FlushFinalBlock()
        Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
        Return encoding.GetString(ms.ToArray())
    End Function
End Class
