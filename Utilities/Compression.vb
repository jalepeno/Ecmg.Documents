'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.IO
Imports System.IO.Compression
Imports System.Text

Namespace Utilities

  Public Class Compression

#Region "Class Constants"

    Private Const BlockSize As Integer = 4096

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Compresses a string into a byte array
    ''' </summary>
    ''' <param name="originalText"></param>
    ''' <returns>Compressed byte array</returns>
    ''' <remarks></remarks>
    ''' 
    Public Shared Function CompressString(ByVal originalText As String) As Byte()
      ' ----- Generate a compressed version of a string
      Try
        '         First, convert the string to a byte array.
        Dim workBytes() As Byte = Encoding.UTF8.GetBytes(originalText)

        ' Write the original length to the end of the array
        ''Dim lintOriginalLength As Integer = workBytes.Length
        'ReDim Preserve workBytes(lintOriginalLength + 4)

        ' ----- Bytes will flow through a memory stream.
        Dim memoryStream As New MemoryStream

        ' ----- Use the newly created memory stream for the compressed data
        Dim zipStream As New GZipStream(memoryStream, CompressionMode.Compress, True)
        With zipStream
          .Write(workBytes, 0, workBytes.Length)
          '.Write(BitConverter.GetBytes(workBytes.Length), workBytes.Length - 1, 4)
          .Flush()
          ' ----- Close the compression stream.
          .Close()
        End With

        ' ----- Return the compressed bytes.
        'Return memoryStream.ToArray

        Dim returnBytes() As Byte = memoryStream.ToArray
        ''Dim lintOriginalReturnLength As Integer = returnBytes.Length
        ''ReDim Preserve returnBytes(returnBytes.Length + 4)

        ''Dim lengthBytes() As Byte = BitConverter.GetBytes(lintOriginalLength)

        ''For i As Integer = 1 To 4
        ''  returnBytes(lintOriginalReturnLength + i) = lengthBytes(i - 1)
        ''Next

        Return returnBytes

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.LogException(ex, String.Format("Compression::CompressString('{0}')", originalText))
        Return Nothing
      End Try

    End Function

    ''' <summary>
    ''' Decompresses a byte array into a string
    ''' </summary>
    ''' <param name="compressed"></param>
    ''' <returns>An uncompressed string</returns>
    ''' <remarks></remarks>
    Public Shared Function DecompressBytes(ByVal compressed As Byte()) As String

      ' ----- Uncompress a previously compressed string.
      Try
        '         Extract the length for a compressed string.
        Dim lastFour(3) As Byte
        Array.Copy(compressed, compressed.Length - 4, lastFour, 0, 4)
        Dim bufferLength As Integer = BitConverter.ToInt32(lastFour, 0)

        ' ----- Create an uncompressed bytes buffer.
        Dim buffer(bufferLength - 1) As Byte

        ' ----- Bytes will flow through a memory stream.
        Dim memoryStream As New MemoryStream(compressed)

        ' ----- Create the decompression stream.
        Dim decompressedStream As New GZipStream(memoryStream, CompressionMode.Decompress, True)

        ' ----- Read and decompress the data into the buffer.
        decompressedStream.Read(buffer, 0, bufferLength)

        ' ----- Convert the bytes into a string.
        Return Encoding.UTF8.GetString(buffer)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return ""
      End Try

    End Function

    ''' <summary>
    ''' Compresses a file
    ''' </summary>
    ''' <param name="sourceFile">The path to a file to compress.</param>
    ''' <param name="destinationFile">The path to write the compressed file to.</param>
    ''' <remarks></remarks>
    Public Shared Sub CompressFile(ByVal sourceFile As String, ByVal destinationFile As String)

      ' ----- Compress the entire contents of a file, and store it in a new file.

      Dim sourcestream As FileStream = Nothing
      Dim destinationStream As FileStream = Nothing
      Dim compressedStream As GZipStream = Nothing

      Try
        '         First, create the input file stream.
        sourcestream = New FileStream(sourceFile, FileMode.Open, FileAccess.Read)

        ' ----- Create the output file stream.
        destinationStream = New FileStream(destinationFile, FileMode.Create, FileAccess.Write)

        ' ----- Bytes will be processed by a compression stream.
        compressedStream = New GZipStream(destinationStream, CompressionMode.Compress, True)

        ' ----- Process bytes from one file into the other.
        Dim buffer(BlockSize) As Byte
        Dim bytesRead As Integer

        Do
          bytesRead = sourcestream.Read(buffer, 0, BlockSize)
          If (bytesRead = 0) Then Exit Do
          compressedStream.Write(buffer, 0, bytesRead)
        Loop

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Compression::CompressFile('{0}', '{1}')", sourceFile, destinationFile))
      Finally
        ' ----- Close all the streams.
        sourcestream.Close()
        compressedStream.Close()
        destinationStream.Close()
      End Try

    End Sub

    ''' <summary>
    ''' Decompresses a file
    ''' </summary>
    ''' <param name="sourceFile">The path to a file to decompress.</param>
    ''' <param name="destinationFile">The path to write the decompressed file to.</param>
    ''' <remarks></remarks>
    Public Shared Sub DecompressFile(ByVal sourceFile As String, ByVal destinationFile As String)

      ' ----- Decompress the entire contents of a file, and store it in a new file.
      Dim sourceStream As FileStream = Nothing
      Dim destinationStream As FileStream = Nothing
      Dim decompressedStream As GZipStream = Nothing

      Try
        '         First, get the files as streams.
        sourceStream = New FileStream(sourceFile, FileMode.Open, FileAccess.Read)
        destinationStream = New FileStream(destinationFile, FileMode.Create, FileAccess.Write)

        ' ----- Bytes will be processed through a decompression stream.
        decompressedStream = New GZipStream(sourceStream, CompressionMode.Decompress, True)

        ' ----- Process bytes from one file into the other.
        Dim buffer(BlockSize) As Byte
        Dim bytesRead As Integer

        Do
          bytesRead = decompressedStream.Read(buffer, 0, BlockSize)
          If (bytesRead = 0) Then Exit Do
          destinationStream.Write(buffer, 0, bytesRead)
        Loop
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Compression::DecompressFile('{0}', '{1}')", sourceFile, destinationFile))
      Finally
        ' ----- Close all the streams
        sourceStream.Close()
        decompressedStream.Close()
        destinationStream.Close()
      End Try

    End Sub

#End Region

    Public Class StringCompression

#Region "Public Shared Methods"

      Public Shared Function CompressString(ByVal lpText As String) As String
        Try

          Dim buffer As Byte() = Encoding.UTF8.GetBytes(lpText)
          Dim lobjMemoryStream As New MemoryStream
          Using lobjZipStream As New GZipStream(lobjMemoryStream, CompressionMode.Compress, True)
            lobjZipStream.Write(buffer, 0, buffer.Length)
          End Using

          lobjMemoryStream.Position = 0

          Dim compressed(lobjMemoryStream.Length) As Byte
          lobjMemoryStream.Read(compressed, 0, compressed.Length)

          Dim gzBuffer(compressed.Length + 4) As Byte

          System.Buffer.BlockCopy(compressed, 0, gzBuffer, 4, compressed.Length)
          System.Buffer.BlockCopy(BitConverter.GetBytes(buffer.Length), 0, gzBuffer, 0, 4)
          Return Convert.ToBase64String(gzBuffer)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Function

      Public Shared Function DecompressString(ByVal lpCompressedText As String) As String
        Try

          Dim gzBuffer() As Byte = Convert.FromBase64String(lpCompressedText)

          Using lobjMemoryStream As New MemoryStream
            Dim lintMsgLength As Integer = BitConverter.ToInt32(gzBuffer, 0)
            lobjMemoryStream.Write(gzBuffer, 4, gzBuffer.Length - 4)

            Dim buffer(lintMsgLength) As Byte

            lobjMemoryStream.Position = 0

            Using lobjZipStream As New GZipStream(lobjMemoryStream, CompressionMode.Decompress)
              lobjZipStream.Read(buffer, 0, buffer.Length)
            End Using

            Return Encoding.UTF8.GetString(buffer)

          End Using

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Function

#End Region

    End Class


  End Class

End Namespace