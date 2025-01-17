'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Security.Cryptography
Imports System.Text
Imports Documents.Utilities

Namespace Encryption

  'Module Module1

  ' <summary>
  ' This class generates and compares hashes using MD5, SHA1, SHA256, SHA384, 
  ' and SHA512 hashing algorithms. Before computing a hash, it appends a
  ' randomly generated salt to the plain text, and stores this salt appended
  ' to the result. To verify another plain text value against the given hash,
  ' this class will retrieve the salt value from the hash string and use it
  ' when computing a new hash of the plain text. Appending a salt value to
  ' the hash may not be the most efficient approach, so when using hashes in
  ' a real-life application, you may choose to store them separately. You may
  ' also opt to keep results as byte arrays instead of converting them into
  ' base64-encoded strings.
  ' </summary>
  Public Class SimpleHash

    Public Shared Function ComputeBasicNumber(ByVal plainText As String) _
                           As Integer
      Try
        Dim hash() As Byte = GetHash(plainText, "MD5", GetHashSalt)
        Dim hashNumber As Integer = BitConverter.ToInt32(hash, 0)

        Return hashNumber

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ComputeHash(ByVal plainText As String) _
                               As String
      Try
        Return ComputeHash(plainText, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ComputeHash(ByVal plainText As String,
                                           ByVal hashAlgorithm As String) _
                               As String
      Try
        Return ComputeHash(plainText, hashAlgorithm, GetHashSalt)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Shared Function ComputeHash(ByVal plainText As String,
                                       ByVal hashAlgorithm As String,
                                       ByVal saltSize As Integer) _
                           As String

      Dim saltBytes() As Byte

      Try
        ' Allocate a byte array, which will hold the salt.
        saltBytes = New Byte(saltSize - 1) {}

        ' Initialize a random number generator.
        Dim rng As RNGCryptoServiceProvider
        rng = New RNGCryptoServiceProvider()

        ' Fill the salt with cryptographically strong byte values.
        rng.GetNonZeroBytes(saltBytes)

        Return ComputeHash(plainText, hashAlgorithm, saltBytes)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <summary>
    ' Generates a hash for the given plain text value and returns a
    ' base64-encoded result. Before the hash is computed, a random salt
    ' is generated and appended to the plain text. This salt is stored at
    ' the end of the hash value, so it can be used later for hash
    ' verification.
    ' </summary>
    ' <param name="plainText">
    ' Plaintext value to be hashed. The function does not check whether
    ' this parameter is null.
    ' </param>
    ' < name="hashAlgorithm">
    ' Name of the hash algorithm. Allowed values are: "MD5", "SHA1",
    ' "SHA256", "SHA384", and "SHA512" (if any other value is specified
    ' MD5 hashing algorithm will be used). This value is case-insensitive.
    ' </param>
    ' < name="saltBytes">
    ' Salt bytes. This parameter can be null, in which case a random salt
    ' value will be generated.
    ' </param>
    ' <returns>
    ' Hash value formatted as a base64-encoded string.
    ' </returns>
    Public Shared Function ComputeHash(ByVal plainText As String,
                                       ByVal hashAlgorithm As String,
                                       ByVal saltBytes() As Byte) _
                           As String

      Try

        Dim hash() As Byte = GetHash(plainText, hashAlgorithm, saltBytes)
        ' Convert result into a base64-encoded string.
        Dim hashValue As String
        hashValue = Convert.ToBase64String(hash)

        ' Return the result.
        Return hashValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <summary>
    ' Generates a hash for the given plain text value and returns a
    ' byte array result result. Before the hash is computed, a random salt
    ' is generated and appended to the plain text. This salt is stored at
    ' the end of the hash value, so it can be used later for hash
    ' verification.
    ' </summary>
    ' <param name="plainText">
    ' Plaintext value to be hashed. The function does not check whether
    ' this parameter is null.
    ' </param>
    ' < name="hashAlgorithm">
    ' Name of the hash algorithm. Allowed values are: "MD5", "SHA1",
    ' "SHA256", "SHA384", and "SHA512" (if any other value is specified
    ' MD5 hashing algorithm will be used). This value is case-insensitive.
    ' </param>
    ' < name="saltBytes">
    ' Salt bytes. This parameter can be null, in which case a random salt
    ' value will be generated.
    ' </param>
    ' <returns>
    ' Hash value formatted as byte array.
    ' </returns>
    Private Shared Function GetHash(ByVal plainText As String,
                                       ByVal hashAlgorithm As String,
                                       ByVal saltBytes() As Byte) _
                           As Byte()
      Try
        ' If salt is not specified, generate it on the fly.
        If (saltBytes Is Nothing) Then

          ' Define min and max salt sizes.
          Dim minSaltSize As Integer
          Dim maxSaltSize As Integer

          minSaltSize = 4
          maxSaltSize = 8

          ' Generate a random number for the size of the salt.
          Dim random As Random
          random = New Random()

          Dim saltSize As Integer
          saltSize = random.Next(minSaltSize, maxSaltSize)

          ' Allocate a byte array, which will hold the salt.
          saltBytes = New Byte(saltSize - 1) {}

          ' Initialize a random number generator.
          Dim rng As RNGCryptoServiceProvider
          rng = New RNGCryptoServiceProvider()

          ' Fill the salt with cryptographically strong byte values.
          rng.GetNonZeroBytes(saltBytes)
        End If

        ' Convert plain text into a byte array.
        Dim plainTextBytes As Byte()
        plainTextBytes = Encoding.UTF8.GetBytes(plainText)

        ' Allocate array, which will hold plain text and salt.
        Dim plainTextWithSaltBytes() As Byte =
            New Byte(plainTextBytes.Length + saltBytes.Length - 1) {}

        ' Copy plain text bytes into resulting array.
        Dim I As Integer
        For I = 0 To plainTextBytes.Length - 1
          plainTextWithSaltBytes(I) = plainTextBytes(I)
        Next I

        ' Append salt bytes to the resulting array.
        For I = 0 To saltBytes.Length - 1
          plainTextWithSaltBytes(plainTextBytes.Length + I) = saltBytes(I)
        Next I

        ' Because we support multiple hashing algorithms, we must define
        ' hash object as a common (abstract) base class. We will specify the
        ' actual hashing algorithm class later during object creation.
        Dim hash As HashAlgorithm

        ' Make sure hashing algorithm name is specified.
        If (hashAlgorithm Is Nothing) Then
          hashAlgorithm = ""
        End If

        ' Initialize appropriate hashing algorithm class.
        Select Case hashAlgorithm.ToUpper()

          Case "MD5"
            hash = New MD5CryptoServiceProvider()

          Case "SHA1"
            hash = New SHA1Managed()

          Case "SHA256"
            hash = New SHA256Managed()

          Case "SHA384"
            hash = New SHA384Managed()

          Case "SHA512"
            hash = New SHA512Managed()

          Case Else
            hash = New MD5CryptoServiceProvider()

        End Select

        ' Compute hash value of our plain text with appended salt.
        Dim hashBytes As Byte()
        hashBytes = hash.ComputeHash(plainTextWithSaltBytes)

        ' Create array which will hold hash and original salt bytes.
        Dim hashWithSaltBytes() As Byte =
                                   New Byte(hashBytes.Length +
                                            saltBytes.Length - 1) {}

        ' Copy hash bytes into resulting array.
        For I = 0 To hashBytes.Length - 1
          hashWithSaltBytes(I) = hashBytes(I)
        Next I

        ' Append salt bytes to the result.
        For I = 0 To saltBytes.Length - 1
          hashWithSaltBytes(hashBytes.Length + I) = saltBytes(I)
        Next I

        Return hashWithSaltBytes

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetHashSalt() As Byte()
      Try
        Return New Byte(7) {Asc("E"), Asc("c"), Asc("m"), Asc("g"), Asc("."), Asc("C"), Asc("t"), Asc("s")}
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <summary>
    ' Compares a hash of the specified plain text value to a given hash
    ' value. Plain text is hashed with the same salt value as the original
    ' hash.
    ' </summary>
    ' <param name="plainText">
    ' Plain text to be verified against the specified hash. The function
    ' does not check whether this parameter is null.
    ' </param>
    ' < name="hashAlgorithm">
    ' Name of the hash algorithm. Allowed values are: "MD5", "SHA1",
    ' "SHA256", "SHA384", and "SHA512" (if any other value is specified
    ' MD5 hashing algorithm will be used). This value is case-insensitive.
    ' </param>
    ' < name="hashValue">
    ' Base64-encoded hash value produced by ComputeHash function. This value
    ' includes the original salt appended to it.
    ' </param>
    ' <returns>
    ' If computed hash mathes the specified hash the function the return
    ' value is true; otherwise, the function returns false.
    ' </returns>
    Public Shared Function VerifyHash(ByVal plainText As String,
                                      ByVal hashAlgorithm As String,
                                      ByVal hashValue As String) _
                           As Boolean

      Try
        ' Convert base64-encoded hash value into a byte array.
        Dim hashWithSaltBytes As Byte()
        hashWithSaltBytes = Convert.FromBase64String(hashValue)

        ' We must know size of hash (without salt).
        Dim hashSizeInBits As Integer
        Dim hashSizeInBytes As Integer

        ' Make sure that hashing algorithm name is specified.
        If (hashAlgorithm Is Nothing) Then
          hashAlgorithm = ""
        End If

        ' Size of hash is based on the specified algorithm.
        Select Case hashAlgorithm.ToUpper()

          Case "SHA1"
            hashSizeInBits = 160

          Case "SHA256"
            hashSizeInBits = 256

          Case "SHA384"
            hashSizeInBits = 384

          Case "SHA512"
            hashSizeInBits = 512

          Case Else ' Must be MD5
            hashSizeInBits = 128

        End Select

        ' Convert size of hash from bits to bytes.
        hashSizeInBytes = hashSizeInBits / 8

        ' Make sure that the specified hash value is long enough.
        If (hashWithSaltBytes.Length < hashSizeInBytes) Then
          Return False
        End If

        ' Allocate array to hold original salt bytes retrieved from hash.
        Dim saltBytes() As Byte = New Byte(hashWithSaltBytes.Length -
                                           hashSizeInBytes - 1) {}

        ' Copy salt from the end of the hash to the new array.
        Dim I As Integer
        For I = 0 To saltBytes.Length - 1
          saltBytes(I) = hashWithSaltBytes(hashSizeInBytes + I)
        Next I

        ' Compute a new hash string.
        Dim expectedHashString As String
        expectedHashString = ComputeHash(plainText, hashAlgorithm, saltBytes)

        ' If the computed hash matches the specified hash,
        ' the plain text value must be correct.
        Return (hashValue = expectedHashString)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

  End Class

  '' <summary>
  '' The main entry point for the application.
  '' </summary>
  'Sub Main()

  '  Dim password As String    ' original password
  '  Dim wrongPassword As String    ' wrong password

  '  password = "myP@5sw0rd"
  '  wrongPassword = "password"

  '  Dim passwordHashMD5 As String
  '  Dim passwordHashSha1 As String
  '  Dim passwordHashSha256 As String
  '  Dim passwordHashSha384 As String
  '  Dim passwordHashSha512 As String

  '  passwordHashMD5 = _
  '         SimpleHash.ComputeHash(password, "MD5", Nothing)
  '  passwordHashSha1 = _
  '         SimpleHash.ComputeHash(password, "SHA1", Nothing)
  '  passwordHashSha256 = _
  '         SimpleHash.ComputeHash(password, "SHA256", Nothing)
  '  passwordHashSha384 = _
  '         SimpleHash.ComputeHash(password, "SHA384", Nothing)
  '  passwordHashSha512 = _
  '         SimpleHash.ComputeHash(password, "SHA512", Nothing)

  '  Console.WriteLine("COMPUTING HASH VALUES")
  '  Console.WriteLine("")
  '  Console.WriteLine("MD5   : {0}", passwordHashMD5)
  '  Console.WriteLine("SHA1  : {0}", passwordHashSha1)
  '  Console.WriteLine("SHA256: {0}", passwordHashSha256)
  '  Console.WriteLine("SHA384: {0}", passwordHashSha384)
  '  Console.WriteLine("SHA512: {0}", passwordHashSha512)
  '  Console.WriteLine("")

  '  Console.WriteLine("COMPARING PASSWORD HASHES")
  '  Console.WriteLine("")
  '  Console.WriteLine("MD5    (good): {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      password, "MD5", _
  '                      passwordHashMD5).ToString())
  '  Console.WriteLine("MD5    (bad) : {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      wrongPassword, "MD5", _
  '                      passwordHashMD5).ToString())
  '  Console.WriteLine("SHA1   (good): {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      password, "SHA1", _
  '                      passwordHashSha1).ToString())
  '  Console.WriteLine("SHA1   (bad) : {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      wrongPassword, "SHA1", _
  '                      passwordHashSha1).ToString())
  '  Console.WriteLine("SHA256 (good): {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      password, "SHA256", _
  '                      passwordHashSha256).ToString())
  '  Console.WriteLine("SHA256 (bad) : {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      wrongPassword, "SHA256", _
  '                      passwordHashSha256).ToString())
  '  Console.WriteLine("SHA384 (good): {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      password, "SHA384", _
  '                      passwordHashSha384).ToString())
  '  Console.WriteLine("SHA384 (bad) : {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      wrongPassword, "SHA384", _
  '                      passwordHashSha384).ToString())
  '  Console.WriteLine("SHA512 (good): {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      password, "SHA512", _
  '                      passwordHashSha512).ToString())
  '  Console.WriteLine("SHA512 (bad) : {0}", _
  '                      SimpleHash.VerifyHash( _
  '                      wrongPassword, "SHA512", _
  '                      passwordHashSha512).ToString())
  'End Sub

  'End Module
  '
End Namespace

' END OF FILE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
