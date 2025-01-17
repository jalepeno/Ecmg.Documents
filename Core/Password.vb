'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  ''' <exclude/>
  <Microsoft.VisualBasic.HideModuleName()>
  Friend Class Password

#Region "Class Constants"

    Private Const PROVIDER As Encryption.Symmetric.Provider = Encryption.Symmetric.Provider.TripleDES
    Private Const KEY As String = "IVEBJBMB"

#End Region

#Region "Class Variables"

    'Private mstrUnEncrypted As String = ""
    'Private mstrEncrypted As String = ""

#End Region

#Region "Constructors"

#End Region

#Region "Friend Methods"

    Public Shared Function Encrypt(ByVal lpValue As String) As Encryption.Data
      Try
        Dim lobjSymmetricEncryption As New Encryption.Symmetric(PROVIDER)
        lobjSymmetricEncryption.Key.Text = KEY

        Dim lobjEncryptedData As Encryption.Data
        lobjEncryptedData = lobjSymmetricEncryption.Encrypt(New Encryption.Data(lpValue))

        Return lobjEncryptedData
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Friend Shared Function Encrypt(ByVal lpBytes() As Byte) As Encryption.Data
      Try
        Dim lobjSymmetricEncryption As New Encryption.Symmetric(PROVIDER)
        lobjSymmetricEncryption.Key.Text = KEY

        Dim lobjEncryptedData As Encryption.Data
        lobjEncryptedData = lobjSymmetricEncryption.Encrypt(New Encryption.Data(lpBytes))

        Return lobjEncryptedData
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Friend Shared Function EncryptToText(ByVal lpValue As String) As String
      'Dim lobjSymmetricEncryption As New Symmetric(PROVIDER)
      'lobjSymmetricEncryption.Key.Text = KEY

      'Dim lobjEncryptedData As Utilities.Encryption.Data
      'lobjEncryptedData = lobjSymmetricEncryption.Encrypt(New Utilities.Encryption.Data(lpValue))

      'Return lobjEncryptedData.Text
      Try
        Return Encrypt(lpValue).Text
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Friend Shared Function EscapedEncrypt(ByVal lpValue As String) As String
      Try
        Return Uri.EscapeDataString(EncryptToText(lpValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function EscapedDecrypt(ByVal lpValue As String) As String
      Try
        Return DecryptToText(Uri.UnescapeDataString(lpValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function EscapedEncryptFromHex(ByVal lpValue As String) As String
      Try
        Return Uri.EscapeDataString(Encrypt(lpValue).Hex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function EscapedDecryptFromHex(ByVal lpValue As String) As String
      Return DecryptFromHex(lpValue).Text
    End Function

    Friend Shared Function DecryptFromText(ByVal lpValue As String) As Encryption.Data
      Try
        Dim decryptedData As Encryption.Data
        Dim sym2 As New Encryption.Symmetric(PROVIDER)
        sym2.Key.Text = KEY
        decryptedData = sym2.Decrypt(New Encryption.Data(lpValue))
        Return decryptedData
      Catch FormatEx As FormatException
        ' Do not log here, simply re-throw the exception
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <remarks>If the parameter is not hex encoded then it is returned unchanged.</remarks>
    Public Shared Function DecryptStringFromHex(ByVal lpValue As String) As String
      Try
        Return DecryptFromHex(lpValue).Text
      Catch FormatEx As FormatException
        Return lpValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function DecryptFromHex(ByVal lpValue As String) As Encryption.Data
      Try
        Dim decryptedData As Encryption.Data
        Dim sym2 As New Encryption.Symmetric(PROVIDER)
        sym2.Key.Text = KEY
        Dim encryptedData As New Encryption.Data
        encryptedData.Hex = lpValue
        decryptedData = sym2.Decrypt(encryptedData)
        Return decryptedData
      Catch FormatEx As FormatException
        ' Do not log here, simply re-throw the exception
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function DecryptFromHex(ByVal lpValue As Byte()) As Encryption.Data
      Try
        Dim decryptedData As Encryption.Data
        Dim sym2 As New Encryption.Symmetric(PROVIDER)
        sym2.Key.Text = KEY
        Dim encryptedData As New Encryption.Data
        encryptedData.Bytes = lpValue
        decryptedData = sym2.Decrypt(encryptedData)
        Return decryptedData
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function DecryptToText(ByVal lpValue As String) As String

      'Dim decryptedData As Utilities.Encryption.Data
      'Dim sym2 As New Utilities.Encryption.Symmetric(PROVIDER)
      'sym2.Key.Text = KEY
      'decryptedData = sym2.Decrypt(New Utilities.Encryption.Data(lpValue))

      'Return decryptedData.Text

      Try
        Return DecryptFromText(lpValue).Text
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

  End Class

End Namespace