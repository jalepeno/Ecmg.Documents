'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Configuration
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core.Content
Imports Documents.Exceptions
Imports Documents.Utilities
Imports Ionic.Zip

#End Region

Namespace Core
  Public Class ContentData
    Implements IDisposable

#Region "Class Variables"

    Private mstrContentPath As String
    Private mobjStorageType As Content.StorageTypeEnum
    Private mlngLength As Long
    Private mobjMemoryStream As Stream
    Private mobjContent As Content
    Private mblnUncompressComplete As Boolean = False

#End Region

#Region "Public Properties"

    Public ReadOnly Property CanRead As Boolean
      Get
        If mobjMemoryStream IsNot Nothing Then
          Return mobjMemoryStream.CanRead
        Else
          Return False
        End If
      End Get
    End Property

    Public ReadOnly Property Content() As Content
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          Return mobjContent
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    Public ReadOnly Property StreamType() As String
      Get
        Try
          If mobjMemoryStream IsNot Nothing Then
            Return mobjMemoryStream.GetType.Name
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    Public Property StorageType() As Content.StorageTypeEnum
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          Return mobjStorageType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Content.StorageTypeEnum)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          mobjStorageType = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlIgnore()>
    Public Property ContentPath() As String
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          Return mstrContentPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          mstrContentPath = value
          If Content.Protocol = Network.ProtocolEnum.file Then
            If File.Exists(mstrContentPath) Then
              SetFileLength()
#If Not SILVERLIGHT = 1 Then
              SetStream(mstrContentPath)
#End If
            Else
              If Not Helper.CallStackContainsMethodName("Rename") Then
                mobjMemoryStream = Nothing 'invalidate since the path may have changed
              End If
            End If
          Else
            ' Don't try to get the content yet, we still need to test that part.
            mobjMemoryStream = Nothing
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Size in bytes of the content stream
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Length() As Long
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          Return mlngLength
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
    End Sub

    Public Sub New(ByVal lpContent As Content)
      Try
        SetContent(lpContent)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Public Setter Methods"

    Public Sub SetContent(ByVal lpContent As Content)
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mobjContent = lpContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Sub SetStream(ByVal lpBytes As Byte())
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If (mobjMemoryStream Is Nothing) Then
          'If lpBytes.Length * 1024 * 1024 < ConnectionSettings.Instance.MaximumInMemoryDocumentMegabytes Then
          Dim lintMegabytes As Integer = lpBytes.Length / 1024 / 1024
          If lintMegabytes < ConfigurationManager.AppSettings.Item("MaximumInMemoryDocumentMegabytes") Then
            mobjMemoryStream = New IO.MemoryStream()
          Else
            Dim lstrTempFilePath As String = Path.GetTempFileName()
            mobjMemoryStream = New FileStream(lstrTempFilePath, FileMode.Open)
            Me.Content.Document.TempFilePaths.Add(lstrTempFilePath)
            Me.Content.Document.TempFileStreams.Add(mobjMemoryStream)
          End If
        End If


        mobjMemoryStream.Write(lpBytes, 0, lpBytes.Length)

        mobjMemoryStream.Position = 0

        mlngLength = mobjMemoryStream.Length

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

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

    Private Sub SetMemoryStream(ByVal lpStream As Stream)
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If lpStream Is Nothing Then
          Throw New ArgumentNullException(NameOf(lpStream))
        ElseIf lpStream.Length = 0 Then
          Throw New ArgumentException("The specified stream is empty", NameOf(lpStream))
        End If

        If lpStream.Position <> 0 Then
          lpStream.Position = 0
        End If

        mobjMemoryStream = lpStream
        mlngLength = mobjMemoryStream.Length

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub SetStream(ByVal lpStream As IO.Stream)

      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If lpStream Is Nothing Then
          Throw New ArgumentNullException(NameOf(lpStream))
        ElseIf lpStream.CanSeek AndAlso lpStream.Length = 0 Then
          Throw New ZeroLengthContentException("Content has zero length.")
        End If

        If lpStream.CanSeek AndAlso lpStream.Position <> 0 Then
          lpStream.Position = 0
        End If

        If TypeOf (lpStream) Is MemoryStream Then
          SetMemoryStream(lpStream)
        ElseIf TypeOf (lpStream) Is FileStream Then
          SetMemoryStream(lpStream)
        Else
          If (mobjMemoryStream Is Nothing) Then
            mobjMemoryStream = New IO.MemoryStream()
          End If
          Dim buf(1023) As Byte
          Dim numRead As Integer = lpStream.Read(buf, 0, 1024)
          While (numRead > 0)
            mobjMemoryStream.Write(buf, 0, numRead)
            numRead = lpStream.Read(buf, 0, 1024)
          End While
          mlngLength = mobjMemoryStream.Length
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Sub SetFileLength()
      Try

        If IO.File.Exists(mstrContentPath) = False Then
          mlngLength = 0
          Exit Sub
        End If

        Dim lobjFileInfo As New IO.FileInfo(mstrContentPath)
        mlngLength = lobjFileInfo.Length

      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        'Throw
      End Try
    End Sub

#End Region

#Region " IDisposable Support "

    Private disposedValue As Boolean = False    ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: free other state (managed objects).
          If (mobjMemoryStream IsNot Nothing) Then
            mobjMemoryStream.Close()
            mobjMemoryStream.Dispose()
            GC.Collect()
          End If
        End If

        ' DISPOSETODO: free your own state (unmanaged objects).
        ' TODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

#Region "Read Contents Into Stream"

    Friend Function ToStream() As IO.Stream
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If (mobjMemoryStream Is Nothing) Then
          ReadContentsIntoStream()
        End If

        If StorageType = StorageTypeEnum.EncodedCompressed Then
          mobjMemoryStream = UnCompressStream()
        End If

        Return mobjMemoryStream

        'Select Case mobjStorageType
        '  Case StorageTypeEnum.EncodedCompressed
        '    If (mobjMemoryStream Is Nothing) Then
        '      ReadContentsIntoStream()
        '    End If
        '    'Decompress the stream if we are compressed
        '    mobjMemoryStream = UnCompressStream()
        '    Return mobjMemoryStream
        '  Case StorageTypeEnum.EncodedUnCompressed
        '    If (mobjMemoryStream Is Nothing) Then
        '      ReadContentsIntoStream()
        '    End If
        '    Return mobjMemoryStream
        '  Case Else
        '    ' If we already have the stream then just return it, otherwise throw an exception
        '    If (mobjMemoryStream IsNot Nothing) Then
        '      Return mobjMemoryStream
        '    Else
        '      Throw New System.InvalidOperationException(String.Format("Storage type not supported - {0}", mobjStorageType.ToString))
        '    End If
        'End Select

      Catch OutOfMemEx As OutOfMemoryException
        ApplicationLogging.LogException(OutOfMemEx, Reflection.MethodBase.GetCurrentMethod)
        Throw New Exceptions.ContentTooLargeException(Me.Content, OutOfMemEx)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function


    ''' <summary>
    ''' The purpose of the method is to ultimately populate 
    ''' mobjMemoryStream with uncompressed/unencoded data
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReadContentsIntoStream()
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If (mobjMemoryStream IsNot Nothing) AndAlso (mobjMemoryStream.CanRead = True) Then 'already got the stream
          Exit Sub
        End If

        'Read in the ContentPath and chunk it out
        If (mstrContentPath IsNot Nothing) Then
          If (File.Exists(Me.mstrContentPath)) Then

            'Dim lintChunkSize As Integer = 1024
            'Dim lobjBuffer(lintChunkSize - 1) As Byte
            'Dim numBytesRead As Integer = 0
            'Dim lobjMem As New IO.MemoryStream()

            'Using lobjFileStream As FileStream = File.OpenRead(Me.ContentPath)
            '  lobjMem.SetLength(lobjFileStream.Length)
            '  numBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)

            '  While (numBytesRead > 0)
            '    lobjMem.Write(lobjBuffer, 0, numBytesRead)
            '    numBytesRead = lobjFileStream.Read(lobjBuffer, 0, lintChunkSize)
            '  End While
            '  lobjFileStream.Close()
            'End Using
            'lobjMem.Position = 0
            'mobjMemoryStream = lobjMem

            ' <Modified by: Ernie Bahr at 9/6/2012-3:06:55 PM on machine: ERNIEBAHR-THINK>
            mobjMemoryStream = Helper.WriteFileToMemoryStream(Me.ContentPath)
            ' </Modified by: Ernie Bahr at 9/6/2012-3:06:55 PM on machine: ERNIEBAHR-THINK>

            mlngLength = mobjMemoryStream.Length
            'lobjBuffer = Nothing
            Exit Sub
          End If
        End If


        If (mobjContent IsNot Nothing) Then

          Dim settings As New XmlReaderSettings() With {
          .ConformanceLevel = ConformanceLevel.Fragment,
          .IgnoreWhitespace = True,
          .IgnoreComments = True
          }

          Using reader As XmlReader = XmlReader.Create(mobjContent.Version.Document.DeSerializationPath, settings)
            'Figured out a way to get the correct "Data" element because you can have multiple
            'versions with multiple content elements

            'Read ahead until the <version> element
            Do While (reader.Read())
              If reader.Name.Equals("version", StringComparison.CurrentCultureIgnoreCase) Then
                If (reader.GetAttribute("ID") = mobjContent.Version.ID) Then
                  'We found the version, now get the correct content(data)
                  Do While (reader.Read())
                    If reader.Name.Equals("contentpath", StringComparison.CurrentCultureIgnoreCase) Then
                      For Each lobjContent As Content In mobjContent.Version.Contents
                        If (mobjContent.ContentPath.Equals(lobjContent.ContentPath, StringComparison.CurrentCultureIgnoreCase)) Then
                          If (reader.ReadToFollowing("Data") = True) Then
                            reader.ReadStartElement()
                            ReadContentsIntoStream(reader)
                            Exit Sub
                          End If
                        Else
                          'Move to correct content path in the xml file
                          reader.ReadToFollowing("ContentPath")
                        End If
                      Next
                    End If
                  Loop
                End If
              End If

            Loop

          End Using
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub ReadContentsIntoStream(ByVal r As System.Xml.XmlReader)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If r.CanReadBinaryContent Then
          Dim buf(1023) As Byte
          mobjMemoryStream = New IO.MemoryStream()
          Dim numRead As Integer = r.ReadContentAsBase64(buf, 0, 1024)
          While (numRead > 0)
            mobjMemoryStream.Write(buf, 0, numRead)
            numRead = r.ReadContentAsBase64(buf, 0, 1024)
          End While
          buf = Nothing
        Else
          Throw New NotSupportedException()
        End If

        If TypeOf mobjMemoryStream Is MemoryStream Then
          DirectCast(mobjMemoryStream, MemoryStream).Capacity = mobjMemoryStream.Length
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub 'ReadContentsIntoStream

    Private Function UnCompressStream() As IO.Stream
      Try
        If (mobjMemoryStream IsNot Nothing) Then
          If (mblnUncompressComplete = True) Then
            Return mobjMemoryStream
          End If
          mobjMemoryStream.Position = 0
          Select Case mobjStorageType
            Case Content.StorageTypeEnum.EncodedCompressed
              'Read in the zip file
              Dim lobjZipFile As ZipFile = ZipFile.Read(mobjMemoryStream)
              'Set up the output stream
              Dim lobjOutputStream As New IO.MemoryStream '(lobjZipFile.Entries(0).UncompressedSize)
              'Assume 1 entry, get first
              'lobjZipFile.Extract(lobjZipFile.Entries(0).FileName, lobjOutputStream)
              lobjZipFile.Item(0).Extract(lobjOutputStream)
              mblnUncompressComplete = True
              mlngLength = lobjOutputStream.Length
              'mobjMemoryStream.Close() ' close the orginal stream
              'mobjMemoryStream = lobjOutputStream
              Return lobjOutputStream

            Case Content.StorageTypeEnum.EncodedUnCompressed
              Return mobjMemoryStream
            Case Else
              Return mobjMemoryStream
          End Select
        End If
        Return mobjMemoryStream
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function CompressStream(ByVal lpMemStream As IO.MemoryStream) As IO.MemoryStream
      Try

        If mobjStorageType = Content.StorageTypeEnum.EncodedCompressed Then

          Dim lobjMemoryStream As New System.IO.MemoryStream()
          'Using lobjZipFile As ZipFile = New ZipFile(lobjMem)
          Using lobjZipFile As New ZipFile

            'lobjZipFile.AddFileStream("temp.txt", "", lpMemStream)
            lobjZipFile.AddEntry("temp.txt", lpMemStream)
            lobjZipFile.Save(lobjMemoryStream)

            If lobjMemoryStream.CanSeek Then
              lobjMemoryStream.Position = 0
            End If

            mlngLength = lobjMemoryStream.Length

            Return lobjMemoryStream
          End Using
        End If

        mlngLength = lpMemStream.Length
        Return lpMemStream


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region
  End Class
End Namespace
