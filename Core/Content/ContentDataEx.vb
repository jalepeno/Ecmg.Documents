'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Configuration
Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Core.Content
Imports Documents.Utilities

#End Region

Namespace Core

  <Serializable()>
  Partial Public Class ContentData
    Implements IXmlSerializable
    Implements IDisposable

    '#Region "Public Properties"

    '    Public ReadOnly Property Content() As Content
    '      Get
    '        Try
    '          If IsDisposed Then
    '            Throw New ObjectDisposedException(Me.GetType.ToString)
    '          End If
    '          Return mobjContent
    '        Catch ex As Exception
    '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '          '  Re-throw the exception to the caller
    '          Throw
    '        End Try
    '      End Get
    '    End Property

    '    <XmlIgnore()> _
    '    Public Property StorageType() As Content.StorageTypeEnum
    '      Get
    '        Try
    '          If IsDisposed Then
    '            Throw New ObjectDisposedException(Me.GetType.ToString)
    '          End If
    '          Return mobjStorageType
    '        Catch ex As Exception
    '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '          '  Re-throw the exception to the caller
    '          Throw
    '        End Try
    '      End Get
    '      Set(ByVal value As Content.StorageTypeEnum)
    '        Try
    '          If IsDisposed Then
    '            Throw New ObjectDisposedException(Me.GetType.ToString)
    '          End If
    '          mobjStorageType = value
    '        Catch ex As Exception
    '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '          '  Re-throw the exception to the caller
    '          Throw
    '        End Try
    '      End Set
    '    End Property

    '    <XmlIgnore()> _
    '    Public Property ContentPath() As String
    '      Get
    '        Try
    '          If IsDisposed Then
    '            Throw New ObjectDisposedException(Me.GetType.ToString)
    '          End If
    '          Return mstrContentPath
    '        Catch ex As Exception
    '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '          '  Re-throw the exception to the caller
    '          Throw
    '        End Try
    '      End Get
    '      Set(ByVal value As String)
    '        Try
    '          If IsDisposed Then
    '            Throw New ObjectDisposedException(Me.GetType.ToString)
    '          End If
    '          mstrContentPath = value
    '          SetFileLength()
    '          mobjMemoryStream = Nothing 'invalidate since the path may have changed
    '        Catch ex As Exception
    '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '          '  Re-throw the exception to the caller
    '          Throw
    '        End Try
    '      End Set
    '    End Property

    '    ''' <summary>
    '    ''' Size in bytes of the content stream
    '    ''' </summary>
    '    ''' <value></value>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public ReadOnly Property Length() As Long
    '      Get
    '        Try
    '          If IsDisposed Then
    '            Throw New ObjectDisposedException(Me.GetType.ToString)
    '          End If
    '          Return mlngLength
    '        Catch ex As Exception
    '          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '          '  Re-throw the exception to the caller
    '          Throw
    '        End Try
    '      End Get
    '    End Property

    '#End Region

    '#Region "Private Properties"

    '    Private ReadOnly Property IsDisposed() As Boolean
    '      Get
    '        Return disposedValue
    '      End Get
    '    End Property

    '#End Region

    '#Region "Constructors"

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal lpContent As Content)
    '      Try
    '        SetContent(lpContent)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        '  Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '#End Region

    Public Sub SetStream(ByVal lpFilePath As String)
      Try
        Dim lobjFileSize As New FileSize(mlngLength)
        If lobjFileSize.Megabytes < ConfigurationManager.AppSettings("MaximumInMemoryDocumentMegabytes") Then
          SetStream(Helper.WriteFileToMemoryStream(lpFilePath))
        Else
          Dim lobjFileStream As New FileStream(lpFilePath, FileMode.Open)
          If Me.Content.Document IsNot Nothing Then
            Me.Content.Document.TempFileStreams.Add(lobjFileStream)
          Else
            ApplicationLogging.WriteLogEntry("Unable to add file stream to  Document.TempFileStreams, the document reference is not set.", TraceEventType.Warning)
          End If

          SetStream(lobjFileStream)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "Public Setter Methods"

    'Public Sub SetContent(ByVal lpContent As Content)
    '  Try
    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If
    '    mobjContent = lpContent
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Need to work on this because a memory stream created outside of this method
    'throws an exception: "Cannot access internal buffer"
    'Public Sub SetStream(ByVal lpStream As IO.MemoryStream)
    '  mobjMemoryStream.Write(lpStream.GetBuffer, 0, lpStream.Length)
    'End Sub

    'Public Sub SetStream(ByVal lpBytes As Byte())
    '  Try
    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If
    '    If (mobjMemoryStream Is Nothing) Then
    '      mobjMemoryStream = New IO.MemoryStream(lpBytes.Length)
    '    End If
    '    mobjMemoryStream.Write(lpBytes, 0, lpBytes.Length)

    '    mobjMemoryStream.Position = 0

    '    mlngLength = mobjMemoryStream.Length

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Private Sub SetMemoryStream(ByVal lpStream As MemoryStream)
    '  Try
    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    If lpStream Is Nothing Then
    '      Throw New ArgumentNullException("lpStream")
    '    ElseIf lpStream.Length = 0 Then
    '      Throw New ArgumentException("The specified stream is empty", "lpStream")
    '    End If

    '    If lpStream.Position <> 0 Then
    '      lpStream.Position = 0
    '    End If

    '    mobjMemoryStream = lpStream
    '    mlngLength = mobjMemoryStream.Length

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Public Sub SetStream(ByVal lpStream As IO.Stream)

    '  Try
    '    If IsDisposed Then
    '      Throw New ObjectDisposedException(Me.GetType.ToString)
    '    End If

    '    If lpStream Is Nothing Then
    '      Throw New ArgumentNullException("lpStream")
    '    ElseIf lpStream.Length = 0 Then
    '      Throw New ArgumentException("The specified stream is empty", "lpStream")
    '    End If

    '    If lpStream.Position <> 0 Then
    '      lpStream.Position = 0
    '    End If

    '    If TypeOf (lpStream) Is MemoryStream Then
    '      SetMemoryStream(lpStream)
    '    Else
    '      If (mobjMemoryStream Is Nothing) Then
    '        mobjMemoryStream = New IO.MemoryStream()
    '      End If
    '      Dim buf(1023) As Byte
    '      Dim numRead As Integer = lpStream.Read(buf, 0, 1024)
    '      While (numRead > 0)
    '        mobjMemoryStream.Write(buf, 0, numRead)
    '        numRead = lpStream.Read(buf, 0, 1024)
    '      End While
    '      mlngLength = mobjMemoryStream.Length
    '    End If

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Sub

#End Region

#Region "IXmlSerializable Methods"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Read in the contents(Data) to mobjMemoryStream.  Only read it in if we are creating from a stream in document (Deserialize).
    ''' </summary>
    ''' <param name="reader"></param>
    ''' <remarks></remarks>
    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If (Helper.CallStackContainsMethodName("CreateFromStream")) Then
          reader.ReadStartElement()
          ReadContentsIntoStream(reader)
          If mobjMemoryStream.Length > 0 Then
            reader.ReadEndElement()
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Write out the contents(Data) to the cdf file
    ''' </summary>
    ''' <param name="writer"></param>
    ''' <remarks></remarks>
    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        'Only do this if the content is stored encoded
        If (mobjStorageType = Content.StorageTypeEnum.EncodedCompressed OrElse mobjStorageType = Content.StorageTypeEnum.EncodedUnCompressed) Then

          Dim lintChunkSize As Integer = 1024
          Dim lobjBuffer(lintChunkSize - 1) As Byte
          Dim numBytesRead As Integer = 0
          Dim lobjMem As New IO.MemoryStream()

          'Read in the ContentPath and chunk it out
          Try

            If (mobjMemoryStream Is Nothing AndAlso mstrContentPath Is Nothing AndAlso mstrContentPath <> String.Empty) Then
              If (File.Exists(Me.mstrContentPath)) Then
                Using lobjFileStream As FileStream = File.OpenRead(Me.ContentPath)
                  lobjMem.SetLength(lobjFileStream.Length)
                  numBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)

                  While (numBytesRead > 0)
                    lobjMem.Write(lobjBuffer, 0, numBytesRead)
                    numBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)
                  End While


                End Using
                lobjMem = CompressStream(lobjMem)
                writer.WriteBase64(lobjMem.GetBuffer(), 0, lobjMem.Length)

              End If
            Else

              'Populate the mobjMemoryStream
              ReadContentsIntoStream()

              'TEST
              'WriteContentsIntoFile("c:\temp\outfiles\" & mobjContent.FileName)
              'TEST

              If (mobjMemoryStream IsNot Nothing And mobjMemoryStream.Length > 0) Then
                mobjMemoryStream = CompressStream(mobjMemoryStream)
                mobjMemoryStream.Position = 0
                If TypeOf mobjMemoryStream Is MemoryStream Then
                  writer.WriteBase64(DirectCast(mobjMemoryStream, MemoryStream).GetBuffer(), 0, mobjMemoryStream.Length)
                Else
                  Dim lobjBytes(mobjMemoryStream.Length) As Byte
                  writer.WriteBase64(lobjBytes, 0, mobjMemoryStream.Length)
                  'writer.WriteBase64(mobjMemoryStream.GetBuffer(), 0, mobjMemoryStream.Length)
                  mobjMemoryStream.Write(lobjBytes, 0, mobjMemoryStream.Length)
                End If

              End If

            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            '  Re-throw the exception to the caller
            Throw
          End Try

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    ''' <summary>
    '''   Checks for folder paths with redundant backslashes and 
    '''   UNC paths with only a single leading backslash.
    ''' </summary>
    ''' <param name="lpSourcePath" type="String">
    '''     <para>
    '''         The path to check.
    '''     </para>
    ''' </param>
    ''' <remarks>
    '''   This is to clean up paths with too many backslashes or incorrect UNC 
    '''   paths which are often the product of a well intentioned effort to 
    '''   eliminate duplicate backslashes in the middle of a path.
    ''' </remarks>
    ''' <returns>
    '''   If a path with a single leading backslash was provided a clean path 
    '''   with a leading double backslash will be returned, otherwise the 
    '''   cleaned original path will be returned.
    ''' </returns>
    Public Shared Function CleanPath(lpSourcePath As String) As String
      Try
        Dim lobjCleanPathBuilder As New StringBuilder(lpSourcePath.Replace("\\", "\"))
        If lobjCleanPathBuilder.Chars(0) = "\" AndAlso lobjCleanPathBuilder.Chars(1) <> "\" Then
          lobjCleanPathBuilder.Insert(0, "\")
        End If

        Return lobjCleanPathBuilder.ToString()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function ShortenPath(lpSourcePath As String, Optional lpMaxPathLength As Integer = 260) As String
      Try

        If lpSourcePath.Length <= lpMaxPathLength Then
          Return lpSourcePath
        End If

        Dim lstrFolderPath As String = Path.GetDirectoryName(lpSourcePath)
        If lstrFolderPath.Length > lpMaxPathLength Then
          Throw New InvalidOperationException("The folder path is too long to shorten.")
        End If

        Dim lstrFileNameOnly As String = Path.GetFileNameWithoutExtension(lpSourcePath)
        Dim lstrExtension As String = Path.GetExtension(lpSourcePath)

        Dim lintOriginalFileNameLength As Integer = lpSourcePath.Length
        Dim lintFolderPathLength As Integer = lstrFolderPath.Length
        Dim lintFileNameOnlyLength As Integer = lstrFileNameOnly.Length
        Dim lintExtensionLength As Integer = lstrExtension.Length

        Dim lintNewFileNameLength As Integer = lpMaxPathLength - lintFolderPathLength - lintExtensionLength - 1

        If lintNewFileNameLength < 1 Then
          Throw New InvalidOperationException("The file name is too long to shorten.")
        End If

        Dim lstrNewFileName As String = String.Concat(lstrFileNameOnly.AsSpan(0, lintNewFileNameLength), lstrExtension)

        Return Path.Combine(lstrFolderPath, lstrNewFileName)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


#Region "Friend Methods To Retrieve Contents"

    Friend Sub WriteToFile(ByVal lpFileName As String, Optional ByVal lpOverwrite As Boolean = False)

      Try


        Dim lstrNotificationMessageSuffix As String = String.Format(" for version '{0}' of document '{1}'", Me.Content.Version.ID, Me.Content.Version.Document.ID)
        Dim lstrNotificationMessage As String = Nothing

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        lpFileName = Helper.CleanPath(lpFileName)

        If File.Exists(lpFileName) Then
          If (lpOverwrite) Then
            lstrNotificationMessage = String.Format("Deleting existing file '{0}'{1}", lpFileName, lstrNotificationMessageSuffix)
            ApplicationLogging.WriteLogEntry(lstrNotificationMessage, TraceEventType.Information, 54123)
            Debug.WriteLine(lstrNotificationMessage)
            File.Delete(lpFileName)
          Else
            lstrNotificationMessage = String.Format("Skipping existing file '{0}'{1}", lpFileName, lstrNotificationMessageSuffix)
            ApplicationLogging.WriteLogEntry(lstrNotificationMessage, TraceEventType.Information, 54124)
            Debug.WriteLine(lstrNotificationMessage)
            Exit Sub
          End If
        End If

        If lpFileName.Length > 256 Then
          Dim lstrLegalFilePath As String = Helper.ShortenPath(lpFileName, 256)
          ApplicationLogging.LogWarning(String.Format("Filepath '{0}' is too long to write to disk.  Shortening to '{1}'", lpFileName, lstrLegalFilePath), Reflection.MethodBase.GetCurrentMethod)
          lpFileName = lstrLegalFilePath
        End If

        Dim fs As New FileStream(lpFileName, FileMode.CreateNew)
        ReadContentsIntoStream()

        Select Case mobjStorageType
          Case StorageTypeEnum.EncodedCompressed
            'Decompress the stream if we are compressed
            mobjMemoryStream = UnCompressStream()
        End Select

        lstrNotificationMessage = String.Format("Writing file '{0}'{1}", lpFileName, lstrNotificationMessageSuffix)
        ApplicationLogging.WriteLogEntry(lstrNotificationMessage, TraceEventType.Information, 54125)
        Debug.WriteLine(lstrNotificationMessage)

        If TypeOf mobjMemoryStream Is MemoryStream Then
          fs.Write(DirectCast(mobjMemoryStream, MemoryStream).GetBuffer(), 0, mobjMemoryStream.Length)
        Else
          Dim lobjBytes() As Byte = Helper.CopyStreamToByteArray(mobjMemoryStream)
          fs.Write(lobjBytes, 0, mobjMemoryStream.Length)
        End If

        fs.Close()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub 'ReadContentsIntoFile

    Friend Function ToByteArray() As Byte()
      Try

        Dim lobjMemoryStream As MemoryStream = ToStream()
        Return lobjMemoryStream.ToArray

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

#End Region



  End Class

End Namespace