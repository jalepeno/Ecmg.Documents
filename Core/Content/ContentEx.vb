'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Configuration
Imports System.IO
Imports System.Xml.Serialization
Imports Documents.Exceptions
'Imports Ecmg.Meta.Exif
' Imports Ecmg.Meta.Ole
Imports Documents.Files
Imports Documents.Utilities
Imports Ionic.Zip
Imports Newtonsoft.Json

#End Region

Namespace Core

  <Serializable()>
  <XmlInclude(GetType(ReferenceOnlyContent))>
  Partial Public Class Content
    Implements IMetaHolder
    Implements ICloneable

#Region "Class Variables"

    Private mobjAnnotations As New Annotations.Annotations

#End Region

#Region "Public Properties"

    ' ''' <summary>
    ' ''' Gets the current path to the content file
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public ReadOnly Property CurrentPath() As String
    '  Get
    '    Try
    '      Return GetCurrentPath()
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      '  Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property


    ''' <summary>
    ''' Gets a System.IO.FileInfo object based on the ContentPath
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property FileInfo() As FileInfo
      Get
        Try
          Return New FileInfo(ContentPath)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <JsonIgnore()>
    Public ReadOnly Property IsAvailable As Boolean
      Get
        Try
          Return DetermineAvailability()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    '''' <summary>
    '''' Gets the file name of the ContentPath
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property FileName() As String
    '  Get
    '    'Return Path.GetFileName(mstrContentPath)
    '    'Return GetLastPart(mstrContentPath, "\")
    '  End Get
    'End Property



    ''' <summary>
    ''' Gets the mime type for the supplied extension in the registry.
    ''' </summary>
    ''' <returns>A standard mime type name from HKEY_CLASSES_ROOT\.ext where ext is the value passed in as lpExtension.  If the mime type is not found in the registry a default value of text/plain is returned.</returns>
    ''' <remarks></remarks>
    Public Overridable Property MIMEType() As String
      Get
        Try
          If mstrMimeType.Length = 0 Then
            Dim lstrErrorMessage As String = String.Empty
            'mstrMimeType = Helper.GetMIMEType(FileExtension, , lstrErrorMessage)
            mstrMimeType = Analyzer.Instance.GetMimeType(FileExtension, "text/plain")
            If lstrErrorMessage.Length > 0 Then
              mstrMimeType = lstrErrorMessage
            End If
          End If
          Return mstrMimeType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrMimeType = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    '' ''' <summary>
    '' ''' Gets a FileSize object for the contained content element.
    '' ''' </summary>
    '' ''' <value></value>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Public Property FileSize() As FileSize
    ''  Get
    ''    If mobjFileSize Is Nothing Then
    ''      mobjFileSize = New FileSize(FileContentSize)
    ''    End If
    ''    Return mobjFileSize
    ''  End Get
    ''  Set(ByVal Value As FileSize)
    ''    If Helper.CallStackContainsMethodName("Deserialize") Then
    ''      mobjFileSize = Value
    ''    Else
    ''      Throw New InvalidOperationException("Set operation not supported for FileSizeProperty")
    ''    End If
    ''  End Set
    ''End Property


    ''' <summary>
    ''' The unique hash for the content element
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    <JsonIgnore()>
    Public Property Hash() As String
      Get
        If Me.StorageType = StorageTypeEnum.Reference AndAlso
          mstrHash.Length = 0 AndAlso
          Helper.CallStackContainsMethodName("Serialize") Then
          ' Only generate the hash here if it was not already initialized and we are in the process of serialization
          '  Rick is happier now...
          'mstrHash = GenerateHash(mstrContentPath)
          mstrHash = GenerateHash(ContentPath)
        End If
        Return mstrHash
      End Get
      Set(ByVal value As String)
        If Helper.CallStackContainsMethodName("Deserialize") Then
          mstrHash = value
        Else
          Throw New InvalidOperationException("Set operation not supported for Hash Property")
        End If
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the set of annotations 
    ''' for this content element.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Annotations() As Annotations.Annotations
      Get
        Return mobjAnnotations
      End Get
      Set(ByVal value As Annotations.Annotations)
        mobjAnnotations = value
      End Set
    End Property

#End Region

#Region "Methods"

#Region "Public Methods"

#Region "ShouldSerialize Methods"

    Public Function ShouldSerializeContentProperties() As Boolean
      Try
        Return Metadata.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Function ShouldSerializeAnnotations() As Boolean
    '  Try
    '    Return Annotations.Count > 0
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

#End Region

#End Region

#Region "Protected Friend Methods"

    Protected Friend Function GenerateHash(ByVal lpFilePath As String) As String
      Try

        'If lpFilePath = String.Empty Then
        '  ApplicationLogging.WriteLogEntry("Unable to generate hash, the specified file path is empty.", _
        '                                   TraceEventType.Warning, 403)
        '  Return ""
        'Else
        If File.Exists(lpFilePath) = False OrElse lpFilePath = String.Empty Then

          ' Let's first try to get the file
          Try
            Using lobjStream As IO.Stream = Data.ToStream
              Dim lobjBytes(lobjStream.Length) As Byte
              lobjStream.Read(lobjBytes, 0, lobjStream.Length)
              Return HashBytes(lobjBytes)
            End Using

          Catch MemEx As OutOfMemoryException
            ApplicationLogging.LogException(MemEx, Reflection.MethodBase.GetCurrentMethod)
            ApplicationLogging.WriteLogEntry(String.Format(
                                 "Unable to generate hash, the specified file '{0}' is too large.",
                                 lpFilePath), TraceEventType.Warning, 43987)

            Return String.Empty
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ApplicationLogging.WriteLogEntry(String.Format(
                                             "Unable to generate hash, the specified file '{0}' does not exist.",
                                             lpFilePath), TraceEventType.Warning, 404)
            Return String.Empty
          End Try
        Else

          Dim lobjHash As New Encryption.Hash(Encryption.Hash.Provider.SHA256)

          Using lobjStream As New FileStream(lpFilePath, FileMode.Open, FileAccess.Read)
            Dim lobjBytes(lobjStream.Length) As Byte
            lobjStream.Read(lobjBytes, 0, lobjStream.Length)
            Return HashBytes(lobjBytes)
          End Using

        End If


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Friend Shared Function HashBytes(ByVal lpBytesToHash() As Byte) As String
      Try
        Dim lstrHash As String

        ' Calculate the hash and get the text value
        Dim _Hash As New System.Security.Cryptography.SHA256Managed
        Dim b As Byte() = _Hash.ComputeHash(lpBytesToHash)
        lstrHash = System.Text.Encoding.UTF8.GetString(b)
        'mstrHash = lobjHash.Calculate(lobjStream).Text

        ' Encrypt the hash value and store the HEX representation
        lstrHash = Core.Password.Encrypt(lstrHash).Hex

        Return lstrHash
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Friend Methods"

    ''' <summary>
    ''' Deletes the file referenced by the content path (if it exists).
    ''' </summary>
    Friend Sub DeleteContentFiles()
      Try
        If File.Exists(ContentPath) Then
          Try
            'If Version IsNot Nothing AndAlso Version.Document IsNot Nothing AndAlso Version.Document.LogSession IsNot Nothing Then
            '  Version.Document.LogSession.LogVerbose("Deleting file referenced by ContentPath: '{0}'.", ContentPath)
            'End If
            File.Delete(ContentPath)
          Catch ex As Exception
            ApplicationLogging.WriteLogEntry(String.Format("Unable to delete file '{0}' in method DeleteContentFiles: {1}", ContentPath, ex.Message), TraceEventType.Warning)
          End Try
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub


#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Renames the contained content file.
    ''' </summary>
    ''' <param name="lpNewName">The new name to assign without an extension.</param>
    ''' <remarks>The original extension is always retained.</remarks>
    Public Function Rename(ByVal lpNewName As String) As Boolean
      Return Rename(lpNewName, Nothing)
    End Function

    ''' <summary>
    ''' Renames the contained content file.
    ''' </summary>
    ''' <param name="lpNewName">The new name to assign without an extension.</param>
    ''' <param name="lpNewExtension">The new extension to assign.  If an empty string is passed then the extension will not be changed.</param>
    ''' <remarks></remarks>
    Public Function Rename(ByVal lpNewName As String, ByVal lpNewExtension As String) As Boolean

      Dim lstrNewExtension As String

      Try

        ' Get the extension to set for the new file name
        If lpNewExtension IsNot Nothing AndAlso lpNewExtension.Length > 0 Then
          lstrNewExtension = lpNewExtension
        Else
          lstrNewExtension = Me.FileExtension
        End If

        Dim lstrOriginalPath As String = Me.ContentPath
        Dim lstrNewFileName As String = String.Format("{0}.{1}", Helper.CleanFile(Path.GetFileNameWithoutExtension(lpNewName), "_"), lstrNewExtension.TrimStart("."))

        If Not String.IsNullOrEmpty(Me.ContentPath) Then
          Dim lstrNewPath As String = String.Format("{0}\{1}", Path.GetDirectoryName(Me.ContentPath), lstrNewFileName)

          If lstrNewPath = lstrOriginalPath Then
            ' Just skip the move part
          Else
            If File.Exists(lstrNewPath) Then
              ' We do not need to move it, it is already there
              If File.Exists(lstrOriginalPath) Then
                ' This would be replaced in a move, let's delete it.
                File.Delete(lstrOriginalPath)
              End If
            Else
              If (FileInfo IsNot Nothing) AndAlso (FileInfo.Exists) Then
                FileInfo.MoveTo(lstrNewPath)
              End If
            End If
          End If

          Me.ContentPath = lstrNewPath
        Else
          mstrFileName = lstrNewFileName
        End If

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try

    End Function

    Public Shared Function GetStreamFromPath(lpFilePath As String, ByRef lpTempStreamPath As String) As Stream
      Try

        Dim lobjFileInfo As New FileInfo(lpFilePath)
        Dim lintMegabytes = lobjFileInfo.Length / 1024
        If lintMegabytes < ConfigurationManager.AppSettings("MaximumInMemoryDocumentMegabytes") Then ' ConnectionSettings.Instance.MaximumInMemoryDocumentMegabytes Then
          lpTempStreamPath = String.Empty
          Return Helper.WriteFileToMemoryStream(lpFilePath)
        Else
          Dim lstrTempFilePath As String = Path.GetTempFileName()
          lpTempStreamPath = lstrTempFilePath
          Return New FileStream(lstrTempFilePath, FileMode.Open)
          'Me.Content.Document.TempFilePaths.Add(lstrTempFilePath)
          'Me.Content.Document.TempFileStreams.Add(mobjMemoryStream)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ChangeFileName(ByVal lpNewName As String) As Boolean
      Try
        mstrFileName = lpNewName
        Return True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try
    End Function

    Public Shared Function UUEncodeFileToString(ByVal inputFileName As String) As String
      Dim inFile As System.IO.FileStream
      Dim binaryData() As Byte

      Try

        Try
          inFile = New System.IO.FileStream(inputFileName,
                                            System.IO.FileMode.Open,
                                            System.IO.FileAccess.Read)
          ReDim binaryData(inFile.Length)
          Dim bytesRead As Long = inFile.Read(binaryData,
                                              0,
                                              inFile.Length)
          inFile.Close()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try



        ' Convert the binary input into Base64 UUEncoded output.
        Dim base64String As String
        Try
          base64String = System.Convert.ToBase64String(binaryData,
                                                       0,
                                                       binaryData.Length)
        Catch ex As Exception
          Dim lobjConversionException As New Exception("Binary data array is null.", ex)
          ApplicationLogging.LogException(lobjConversionException, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try

        If base64String Is Nothing OrElse base64String.Length = 0 Then
          Throw New Exception(String.Format(
                                            "Unable to UUEncode file '{0}' to string.  The method Convert.ToBase64String returned nothing.  The byte array passed had a length of {1}",
                                            inputFileName, binaryData.Length))
        End If
        ' Returnthe UUEncoded string.
        Return base64String

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub EncodeFile(ByVal lpBytes As Byte(), ByVal lpFileName As String)
      Try
        Dim lobjOutputFile As System.IO.FileStream = EncodeBytesToStream(lpBytes, lpFileName)

        lobjOutputFile.Close()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function EncodeBytesToStream(ByVal lpBytes As Byte(), ByVal lpFileName As String) As IO.Stream
      Try
        Dim lobjOutputFile As System.IO.FileStream

        lobjOutputFile = New System.IO.FileStream(lpFileName,
                                  System.IO.FileMode.Create,
                                  System.IO.FileAccess.Write)

        For lintByteCounter As Integer = 0 To lpBytes.Length - 1
          lobjOutputFile.WriteByte(lpBytes(lintByteCounter))
        Next

        Return lobjOutputFile

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Sub DecodeFileFromString(ByVal lpEncodedData As String, ByVal lpOutputFilePath As String, ByVal lpDecompress As Boolean)
    '  Try

    '    '  Take the encoded file and create a stream object from it
    '    Dim lobjMemoryStream As New IO.MemoryStream(Convert.FromBase64String(lpEncodedData))

    '    Dim lobjFileStream As New FileStream(lpOutputFilePath, FileMode.Create)
    '    Dim lintBufferLength As Integer = 256
    '    Dim lobjBuffer(lintBufferLength) As Byte
    '    Dim lintBytesRead As Integer = lobjMemoryStream.Read(lobjBuffer, 0, lintBufferLength)
    '    '  Write the required bytes
    '    While lintBytesRead > 0
    '      lobjFileStream.Write(lobjBuffer, 0, lintBytesRead)
    '      lintBytesRead = lobjMemoryStream.Read(lobjBuffer, 0, lintBufferLength)
    '    End While

    '    lobjMemoryStream.Close()
    '    lobjFileStream.Close()


    '    If lpDecompress = True Then
    '      Dim lstrTempFilePath As String = lpOutputFilePath & ".zip"
    '      IO.File.Move(lpOutputFilePath, lstrTempFilePath)
    '      Dim lobjZipFile As New ZipFile(lstrTempFilePath)
    '      'lobjZipFile.ExtractAll(IO.Path.GetPathRoot(lpOutputFilePath))
    '      For Each lobjZipEntry As ZipEntry In lobjZipFile
    '        lobjZipEntry.Extract(IO.Path.GetPathRoot(lpOutputFilePath), True)
    '      Next
    '      lobjMemoryStream.Close()

    '    End If


    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Shared Sub DecodeData(ByVal lpFilePath As String, ByVal lpDecompress As Boolean)
      Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Public Shared Function GetXpsProperties(ByVal lpXpsFile As String) As IProperties
    '  Try

    '    Dim lobjEcmProperties As New ECMProperties
    '    Dim lobjXpsDocument As New XpsDocument(lpXpsFile, FileAccess.Read)

    '    Dim lobjProperties As PackageProperties = lobjXpsDocument.CoreDocumentProperties

    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Category", lobjProperties.Category))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Content Status", lobjProperties.ContentStatus))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Content Type", lobjProperties.ContentType))
    '    ' ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDate, "Date Created", lobjProperties.Created))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Author", lobjProperties.Creator))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Comments", lobjProperties.Description))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Identifier", lobjProperties.Identifier))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Keywords", lobjProperties.Keywords))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Language", lobjProperties.Language))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Last Author", lobjProperties.LastModifiedBy))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDate, "Date Last Printed", lobjProperties.LastPrinted))
    '    ' ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDate, "Date Last Modified", lobjProperties.Modified))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Revision", lobjProperties.Revision))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Subject", lobjProperties.Subject))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Title", lobjProperties.Title))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Version", lobjProperties.Version))

    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Category", lobjProperties.Category))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Content Status", lobjProperties.ContentStatus))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Content Type", lobjProperties.ContentType))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDate, "Date Created", lobjProperties.Created))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Author", lobjProperties.Creator))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Comments", lobjProperties.Description))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Identifier", lobjProperties.Identifier))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Keywords", lobjProperties.Keywords))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Language", lobjProperties.Language))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Last Author", lobjProperties.LastModifiedBy))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDate, "Date Last Printed", lobjProperties.LastPrinted))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDate, "Date Last Modified", lobjProperties.Modified))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Revision", lobjProperties.Revision))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Subject", lobjProperties.Subject))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Title", lobjProperties.Title))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Version", lobjProperties.Version))

    '    Return lobjEcmProperties

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Gets all of the EXIF proeprties of the image if available
    '''' </summary>
    '''' <param name="lpImageFile">The fully qualified file name</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function GetExifProperties(ByVal lpImageFile As String) As IProperties
    '  Try
    '    Dim lobjEcmProperties As New ECMProperties
    '    Dim lobjExifData As New ExifData(lpImageFile)

    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Title", lobjExifData.Title))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Artist", lobjExifData.Artist))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Description", lobjExifData.Description))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "EquipmentMaker", lobjExifData.EquipmentMaker))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "EquipmentModel", lobjExifData.EquipmentModel))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDouble, "Aperture", lobjExifData.Aperture))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Copyright", lobjExifData.Copyright))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDate, "DateTimeDigitized", lobjExifData.DateTimeDigitized))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDate, "DateTimeLastModified", lobjExifData.DateTimeLastModified))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDate, "DateTimeOriginal", lobjExifData.DateTimeOriginal))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "ExposureMeteringMode", lobjExifData.ExposureMeteringMode.ToString))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "ExposureProgram", lobjExifData.ExposureProgram.ToString))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDouble, "ExposureTime", lobjExifData.ExposureTime))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "FlashMode", lobjExifData.FlashMode.ToString))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDouble, "FocalLength", lobjExifData.FocalLength))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "FStop", lobjExifData.FStop))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmLong, "Width", lobjExifData.Width))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmLong, "Height", lobjExifData.Height))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDouble, "ResolutionX", lobjExifData.ResolutionX))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDouble, "ResolutionY", lobjExifData.ResolutionY))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmLong, "ISO", lobjExifData.ISO))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "LightSource", lobjExifData.LightSource.ToString))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Orientation", lobjExifData.Orientation.ToString))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Software", lobjExifData.Software))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmDouble, "SubjectDistance", lobjExifData.SubjectDistance))
    '    ''lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "UserComment", lobjExifData.UserComment))

    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Title", lobjExifData.Title))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Artist", lobjExifData.Artist))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Description", lobjExifData.Description))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "EquipmentMaker", lobjExifData.EquipmentMaker))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "EquipmentModel", lobjExifData.EquipmentModel))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDouble, "Aperture", lobjExifData.Aperture))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Copyright", lobjExifData.Copyright))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDate, "DateTimeDigitized", lobjExifData.DateTimeDigitized))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDate, "DateTimeLastModified", lobjExifData.DateTimeLastModified))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDate, "DateTimeOriginal", lobjExifData.DateTimeOriginal))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "ExposureMeteringMode", lobjExifData.ExposureMeteringMode))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "ExposureProgram", lobjExifData.ExposureProgram))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "ExposureTime", lobjExifData.ExposureTime))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "FlashMode", lobjExifData.FlashMode))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDouble, "FocalLength", lobjExifData.Title))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "FStop", lobjExifData.FStop))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmLong, "Width", lobjExifData.Width))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmLong, "Height", lobjExifData.Height))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDouble, "ResolutionX", lobjExifData.ResolutionX))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDouble, "ResolutionY", lobjExifData.ResolutionY))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmLong, "ISO", lobjExifData.ISO))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "LightSource", lobjExifData.LightSource))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Orientation", lobjExifData.Orientation))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Software", lobjExifData.Software))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmDouble, "SubjectDistance", lobjExifData.SubjectDistance))
    '    lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "UserComment", lobjExifData.UserComment))

    '    lobjExifData.Dispose()

    '    Return lobjEcmProperties

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Private Shared Function AddStringProperty(ByVal lpValue As String, ByVal lpPropertyName As String, ByVal lpProperties As ECMProperties) As Boolean
      Try

        If Not String.IsNullOrEmpty(lpValue) Then
          'lpProperties.Add(New ECMProperty(PropertyType.ecmString, lpPropertyName, lpValue.Trim(" ")))
          lpProperties.Add(PropertyFactory.Create(PropertyType.ecmString, lpPropertyName, lpValue.Trim(" ")))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Function GetXmpProperties(ByVal lpImageFile As String) As IProperties
    '  Try

    '    Dim lobjEcmProperties As New ECMProperties

    '    Using lobjFile As New FileStream(lpImageFile, FileMode.Open)
    '      Dim lobjDecoder As New JpegBitmapDecoder(lobjFile, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default)
    '      Dim lobjFrame As BitmapFrame = lobjDecoder.Frames(0)
    '      Dim lobjMetadata As BitmapMetadata = lobjFrame.Metadata

    '      '  Get the single valued image properties

    '      AddStringProperty(lobjMetadata.Title, "Title", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.Subject, "Subject", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.ApplicationName, "ApplicationName", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.CameraManufacturer, "CameraManufacturer", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.CameraModel, "CameraModel", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.Comment, "Comment", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.Copyright, "Copyright", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.DateTaken, "DateTaken", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.Format, "Format", lobjEcmProperties)
    '      AddStringProperty(lobjMetadata.Location, "Location", lobjEcmProperties)
    '      'lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmLong, "Rating", lobjMetadata.Rating))
    '      lobjEcmProperties.Add(PropertyFactory.Create(PropertyType.ecmLong, "Rating", lobjMetadata.Rating))

    '      If lobjMetadata.Author IsNot Nothing Then
    '        'Dim lobjAuthorProperty As New ECMProperty(PropertyType.ecmString, "Authors", Cardinality.ecmMultiValued)
    '        Dim lobjAuthorProperty As MultiValueStringProperty = PropertyFactory.Create(PropertyType.ecmString,
    '                                                               "Authors", "Authors", Cardinality.ecmMultiValued)
    '        For Each lstrAuthor As String In lobjMetadata.Author
    '          lobjAuthorProperty.Values.Add(New Value(lstrAuthor.Trim(" ")))
    '        Next
    '        If lobjAuthorProperty.Values.Count > 0 Then
    '          'lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Authors", lobjAuthorProperty))
    '          lobjEcmProperties.Add(lobjAuthorProperty)
    '        End If
    '      End If

    '      If lobjMetadata.Keywords IsNot Nothing Then
    '        'Dim lobjKeywordsProperty As New ECMProperty(PropertyType.ecmString, "Keywords", Cardinality.ecmMultiValued)
    '        Dim lobjKeywordsProperty As MultiValueStringProperty = PropertyFactory.Create(PropertyType.ecmString,
    '                                                                 "Keywords", "Keywords", Cardinality.ecmMultiValued)
    '        For Each lstrKeyword As String In lobjMetadata.Keywords
    '          lobjKeywordsProperty.Values.Add(New Value(lstrKeyword))
    '        Next
    '        If lobjKeywordsProperty.Values.Count > 0 Then
    '          'lobjEcmProperties.Add(New ECMProperty(PropertyType.ecmString, "Keywords", lobjKeywordsProperty))
    '          lobjEcmProperties.Add(lobjKeywordsProperty)
    '        End If
    '      End If

    '    End Using

    '    Return lobjEcmProperties

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Sub GetAvailableMetadata()
      Try
        ' Add the available properties
        GetAvailableMetadata(ContentPath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Sub

    Private Sub GetAvailableMetadata(ByVal lpContentPath As String)
      Try
        ' Add the available properties
        '  TODO: Add Pdf, Rtf, Autocad
        Select Case IO.Path.GetExtension(lpContentPath).ToLower
          'Case ".doc", ".xls", ".ppt", ".vsd", ".prj", ".pub"
          '  ' The current Ecmg.Meta.Ole implementation only works in 32bit Windows`
          '  If Helper.Is32BitOS Then
          '    Metadata.Add(GetOLEProperties(lpContentPath))
          '  End If
          Case ".jpg", ".jpeg"
            'Metadata.Add(GetExifProperties(lpContentPath))
            'Metadata.Add(GetXmpProperties(lpContentPath))
            'Case ".xps"
            '  Metadata.Add(GetXpsProperties(lpContentPath))
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Provide a warning in the log and move on
        ApplicationLogging.WriteLogEntry(String.Format("An error occured while trying to get available metadata for the file '{0}': {1}",
                                                    lpContentPath, ex.Message), TraceEventType.Warning)
      Finally

      End Try
    End Sub

    'Public Sub SyncronizeVersionProperties(lpContentPropertyAccessor As IContentPropertyAccessor)
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Shared Function TrimFileName(ByVal lpFileName As String) As String
      Try

        Dim lstrTrimmedFileName As String = lpFileName

        ' Make sure the total file path is under 260 characters
        If lpFileName.Length > MAX_FILE_LENGTH Then
          lstrTrimmedFileName = String.Format("{0}.{1}", lstrTrimmedFileName.Substring(0, MAX_FILE_LENGTH),
                                              lstrTrimmedFileName.Substring(lstrTrimmedFileName.LastIndexOf("."c) + 1))
        End If

        Return lstrTrimmedFileName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Try

        If disposedValue = True Then
          Return "Content Disposed"
        End If

        Dim lstrReturnValue As String = String.Empty

        If FileName Is Nothing OrElse FileName.Length = 0 Then
          lstrReturnValue = "No Content"
        Else
          If Data Is Nothing Then
            lstrReturnValue = String.Format("{0} - Data not set", FileName)
          Else
            If FileSize IsNot Nothing Then
              lstrReturnValue = String.Format("{0} - {1}", FileName, FileSize.ToString)
            Else
              lstrReturnValue = String.Format("{0}", FileName)
            End If
          End If

          'If Annotations IsNot Nothing AndAlso Annotations.Count > 0 Then
          '  lstrReturnValue += String.Format(" + {0} annotations", Annotations.Count)
          'End If
        End If

        Return lstrReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub RecreateContentFile()
      Try
        Dim lobjData As Byte() = Convert.FromBase64String(Me.Data.ToString)
        File.WriteAllBytes(ContentPath, lobjData)

        '  Check to see if the content was previously compressed
        If PreviousStorageType = StorageTypeEnum.EncodedCompressed Then
          '  We need to de-compress the content file

          '  Rename the file as a zip file
          Dim lstrZippedFileName As String = String.Format("{0}.zip", ContentPath)
          File.Move(ContentPath, lstrZippedFileName)
          Dim lobjZipFile As New ZipFile(lstrZippedFileName)

          For Each lobjZipEntry As ZipEntry In lobjZipFile
            lobjZipEntry.Extract(IO.Path.GetPathRoot(ContentPath), True)
          Next

          '  Delete the zipped version
          File.Delete(lstrZippedFileName)

        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToByteArray() As Byte()
      Try
        If CanRead AndAlso mobjContentData.Length > 0 Then
          Return mobjContentData.ToByteArray
        Else
          Return Nothing
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Converts the current content to a stream
    ''' </summary>
    ''' <returns>A MemoryStream or a FileStream depending upon the current StorageType</returns>
    ''' <remarks></remarks>
    Public Function ToStream() As IO.Stream
      Try
        Select Case StorageType
          Case StorageTypeEnum.EncodedCompressed, StorageTypeEnum.EncodedUnCompressed
            Return mobjContentData.ToStream()
            '  'Decode the bytes from the data
            '  Dim lobjZipBytes As Byte() = Convert.FromBase64String(Me.Data.ToString)
            '  'Read in the zip file
            '  Dim lobjZipFile As Ionic.Utils.Zip.ZipFile = Ionic.Utils.Zip.ZipFile.Read(lobjZipBytes)
            '  'Set up the output stream
            '  Dim lobjOutputStream As New IO.MemoryStream(lobjZipFile.Entries(0).UncompressedSize)
            '  'Assume 1 entry, get first
            '  lobjZipFile.Extract(lobjZipFile.Entries(0).FileName, lobjOutputStream)
            '  Return lobjOutputStream


            'Case StorageTypeEnum.EncodedUnCompressed
            '  Dim lobjData As Byte() = Convert.FromBase64String(Me.Data.ToString)
            '  Return New IO.MemoryStream(lobjData)



            'Dim lobjFileStream As IO.FileStream = File.OpenRead(Me.ContentPath) 'As New IO.FileStream(Me.ContentPath, FileMode.Create)
            'Return lobjFileStream

          Case Else

            If mobjContentData IsNot Nothing AndAlso mobjContentData.Length AndAlso mobjContentData.CanRead Then
              Return mobjContentData.ToStream
            Else
              'Return New IO.FileStream(Me.ContentPath, FileMode.Create)
              If (Not String.IsNullOrEmpty(Me.ContentPath)) AndAlso (File.Exists(Me.ContentPath)) Then
                Dim lobjFileStream As IO.FileStream = File.OpenRead(Me.ContentPath)
                Return lobjFileStream
              Else
                Throw New Exception(String.Format("Unable to read content data for file - {0}.", Me.FileName))
              End If
            End If

        End Select

      Catch OutOfMemEx As OutOfMemoryException
        ApplicationLogging.LogException(OutOfMemEx, Reflection.MethodBase.GetCurrentMethod)
        Throw New Exceptions.ContentTooLargeException(Me, OutOfMemEx)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Function ToFileStream(ByRef lpTempFilePath As String) As FileStream
    'Try
    '    Dim lstrTempFilePath As String = Path.GetTempFileName()
    '    Dim lobjOutputStream As New FileStream("", FileMode.Create)
    '      lpTempFilePath = lstrTempFilePath
    '      'Return New FileStream(lstrTempFilePath, FileMode.Open)
    '    'lobjOutputStream.wr
    '    Helper.CopyStream()
    'Catch ex As Exception
    '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '  '  Re-throw the exception to the caller
    '  Throw
    'End Try
    'End Function

    ''' <summary>
    ''' Creates and returns an independent memory stream copy of the content.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ToMemoryStream() As MemoryStream
      Try
        Dim lobjOutputStream As New MemoryStream
        Dim lintBytesRead As Integer
        Dim lintTotalBytesRead As Integer
        Dim lintBufferSize As Integer = 4096
        Dim lobjBuffer() As Byte
        ReDim lobjBuffer(lintBufferSize)

        Dim lobjContentStream As IO.Stream = Me.ToStream

        If TypeOf lobjContentStream Is MemoryStream Then
          If lobjContentStream.CanSeek Then
            lobjContentStream.Seek(0, SeekOrigin.Begin)
          End If

          Return lobjContentStream

        Else

          Using lobjContentStream
            Do
              lintBytesRead = lobjContentStream.Read(lobjBuffer, 0, lobjBuffer.Length)
              lintTotalBytesRead += lintBytesRead
              lobjOutputStream.Write(lobjBuffer, 0, lintBytesRead)
            Loop Until lintBytesRead <= 0
          End Using

          If lobjOutputStream.CanSeek Then
            lobjOutputStream.Seek(0, SeekOrigin.Begin)
          End If

          Return lobjOutputStream

        End If

        'Using lobjContentStream As IO.Stream = Me.ToStream

        '  If TypeOf lobjContentStream Is MemoryStream Then
        '    If lobjContentStream.CanSeek Then
        '      lobjContentStream.Seek(0, SeekOrigin.Begin)
        '    End If
        '    Return lobjContentStream
        '  Else

        '    Do
        '      lintBytesRead = lobjContentStream.Read(lobjBuffer, 0, lobjBuffer.Length)
        '      lintTotalBytesRead += lintBytesRead
        '      lobjOutputStream.Write(lobjBuffer, 0, lintBytesRead)
        '    Loop Until lintBytesRead <= 0

        '  End If

        'End Using

        'If lobjOutputStream.CanSeek Then
        '  lobjOutputStream.Seek(0, SeekOrigin.Begin)
        'End If

        'Return lobjOutputStream

      Catch OutOfMemEx As OutOfMemoryException
        ApplicationLogging.LogException(OutOfMemEx, Reflection.MethodBase.GetCurrentMethod)
        Throw New Exceptions.ContentTooLargeException(Me, OutOfMemEx)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToNamedStream() As INamedStream
      Try
        Return New NamedStream(ToMemoryStream, FileName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub WriteToFile(ByVal lpFileName As String, Optional ByVal lpOverwrite As Boolean = False)
      Try
        mobjContentData.WriteToFile(lpFileName, lpOverwrite)
        'Select Case StorageType
        '  Case StorageTypeEnum.EncodedCompressed, StorageTypeEnum.EncodedUnCompressed
        '    mobjContentData.WriteToFile(lpFileName, lpOverwrite)
        '  Case Else
        '    File.Copy(Me.ContentPath, lpFileName, lpOverwrite)
        'End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Takes the content from the specified path and locates it within the Serialized location
    ''' </summary>
    ''' <remarks></remarks>
    Friend Overridable Sub UpdateContentLocation(ByVal lpContentReferenceBehavior As Document.ContentReferenceBehavior)

      Dim lstrVersionPath As String

      Try

        If Me.StorageType <> StorageTypeEnum.Reference Then
          ' We only need to do this if the storage type is reference
          Exit Sub
        End If

        ' Check to see if the ContentPath or the data is available before doing anything else.
        If File.Exists(ContentPath) = False AndAlso Data.Length = 0 Then
          Dim lstrErrorMessage As String = String.Format(
            "Unable to update content location, the content path '{0}' does not point to a valid file and the content data is null.", ContentPath)
          ' See if this is simply a result of a transform operation, it could be that we only had the cdf file in the reader.
          If Helper.CallStackContainsMethodName("Transform") OrElse Helper.CallStackContainsMethodName("UpdateFileWithStyleSheet") Then
            ApplicationLogging.WriteLogEntry(lstrErrorMessage, TraceEventType.Verbose, 62348)
            Exit Sub
          Else
            Throw New DocumentException(Me.Version.Document, Me.Version.ID, lstrErrorMessage)
          End If
        End If

        Select Case lpContentReferenceBehavior
          Case Core.Document.ContentReferenceBehavior.Copy, Core.Document.ContentReferenceBehavior.Move

            If Me.Document IsNot Nothing AndAlso Me.Document.SerializationPath <> String.Empty Then
              ' Move the content to the current serialization path

              ' Cache the file name in case it changes as a result of the operation below
              Dim lstrFileName As String = Me.FileName

              Dim lstrSerializationPath As String = Me.Document.SerializationPath
              'lstrVersionPath = IO.Path.GetDirectoryName(Me.Document.SerializationPath) & "\" & Version.ID
              ' Changed by EFB 1/6/09
              '   IO.Path.GetDirectoryName will return the original long directory name from a short 8.3 file path
              '   We do not want that behavior here
              '   So we will instead use FileHelper.GetDirectoryName which 
              '   resolves the directory name based on string parsing
              lstrVersionPath = String.Format("{0}\{1}", Path.GetDirectoryName(lstrSerializationPath), Version.ID)
              lstrVersionPath = Helper.RemoveExtraBackSlashesFromFilePath(lstrVersionPath)
              Dim lstrNewContentPath As String
              'lstrNewContentPath = String.Format("{0}\{1}", lstrVersionPath, FileName)
              If FileName.Length > 0 Then
                lstrNewContentPath = String.Format("{0}\{1}", lstrVersionPath, FileName)
              Else
                lstrNewContentPath = String.Format("{0}\{1}", lstrVersionPath, Path.GetFileName(ContentPath))
              End If

              If ContentPath <> lstrNewContentPath OrElse (ContentPath IsNot Nothing AndAlso IO.File.Exists(ContentPath) = False) Then

                Try
                  Directory.CreateDirectory(lstrVersionPath)
                Catch PathEx As InvalidOperationException
                  ApplicationLogging.LogException(PathEx, Reflection.MethodBase.GetCurrentMethod)
                  Throw New DocumentException(Document, Version.ID, "Unable to create directory for version '" &
                    Document.ID & ":" & Version.ID & "' Export.", PathEx)
                End Try

                If IO.File.Exists(lstrNewContentPath) Then

                  ' Check to see if the file is readonly, if it is fix it so we can safely delete it
                  If File.GetAttributes(lstrNewContentPath) And FileAttributes.ReadOnly = FileAttributes.ReadOnly Then
                    ' Reset the file to normal
                    File.SetAttributes(lstrNewContentPath, FileAttributes.Normal)
                  End If

                  ' Delete the file
                  IO.File.Delete(lstrNewContentPath)
                End If

                If lpContentReferenceBehavior = Core.Document.ContentReferenceBehavior.Move Then

                  Dim lstrContentDirectory As String = IO.Path.GetDirectoryName(ContentPath)

                  If File.Exists(ContentPath) Then
                    IO.File.Move(ContentPath, lstrNewContentPath)
                  ElseIf Data.Length > 0 Then
                    Data.WriteToFile(lstrNewContentPath, True)
                  Else
                    Throw New DocumentException(Me.Version.Document, Me.Version.ID,
                      String.Format("Unable to update content location, the content path '{0}' does not point to a valid file and the content data is null.", ContentPath))
                  End If

                  ' If there are no more files in the content folder then delete the folder
                  Dim lobjDirectory As New IO.DirectoryInfo(lstrContentDirectory)
                  Dim lobjFileInfo As IO.FileInfo() = lobjDirectory.GetFiles("*.*")

                  If lobjFileInfo.Length = 0 Then
                    Try
                      lobjDirectory.Delete()
                    Catch IoEx As IOException
                      ApplicationLogging.WriteLogEntry(String.Format("Unable to delete content directory '{0}':{1}.", lobjDirectory.FullName, IoEx.Message), TraceEventType.Warning, 63251)
                    End Try

                  End If

                  '   See if there any files in the parent directory
                  '   It should be the directory for the document
                  lobjDirectory = IO.Directory.GetParent(lstrContentDirectory)
                  If (lobjDirectory.GetFiles("*.*").Length = 0) AndAlso (lobjDirectory.GetDirectories.Length = 0) Then
                    Try
                      lobjDirectory.Delete()
                    Catch IoEx As IOException
                      ApplicationLogging.WriteLogEntry(String.Format("Unable to delete parent directory '{0}':{1}.", lobjDirectory.FullName, IoEx.Message), TraceEventType.Warning, 63252)
                    End Try
                  End If

                ElseIf lpContentReferenceBehavior = Core.Document.ContentReferenceBehavior.Copy Then

                  If File.Exists(ContentPath) Then
                    IO.File.Copy(ContentPath, lstrNewContentPath)
                  ElseIf Data.Length > 0 Then
                    Data.WriteToFile(lstrNewContentPath, True)
                  Else
                    Throw New DocumentException(Me.Version.Document, Me.Version.ID,
                      String.Format("Unable to update content location, the content path '{0}' does not point to a valid file and the content data is null.", ContentPath))
                  End If
                End If


                ContentPath = lstrNewContentPath
                mstrFileName = lstrFileName

                ' Set the RelativePath 
                RelativePath = Core.RelativePath.GetRelativePath(lstrSerializationPath, lstrNewContentPath, "\")

              End If

            End If

          Case Else
            ' Do Nothing
            Exit Sub

        End Select


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Sub

    ''' <summary>
    ''' Updates the MIMEType of the current content based 
    ''' on the registered extensions in the Windows Registry.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Overridable Sub UpdateMimeType()
      Try
        UpdateMimeType(String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Updates the MIMEType of the current content
    ''' </summary>
    ''' <param name="lpNewMimeType">The new MIMEType to assign.</param>
    ''' <remarks>
    ''' If no new MIMEType is specified, the mime type will be 
    ''' resolved to the registered extensions in the Windows registry.
    ''' </remarks>
    Friend Overridable Sub UpdateMimeType(ByVal lpNewMimeType As String)
      Try
        If String.IsNullOrEmpty(lpNewMimeType) Then
          MIMEType = Helper.GetMIMEType(Me.FileExtension)
        Else
          MIMEType = lpNewMimeType
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Set's the ContentPath relative to the parent document's SerializationPath
    ''' </summary>
    ''' <remarks></remarks>
    Friend Overridable Sub SetContentPathFromDocumentSerializationPath()
      Try

        Dim lstrFileName As String
        Dim lstrSerializationPath As String
        Dim lstrVersionPath As String

        If Me.FileName Is Nothing OrElse Me.FileName.Length = 0 Then
          Throw New InvalidOperationException("Unable to set content path, the FileName property is not set.")
        End If

        If Me.Document Is Nothing Then
          Throw New InvalidOperationException("Unable to set content path, the Document property is not set.")
        End If

        If Me.Document.SerializationPath.Length = 0 Then
          Throw New InvalidOperationException("Unable to set content path, the document's SerializationPath property is not set.")
        End If

        If FileName.Length > 0 Then
          lstrFileName = FileName
        Else
          'lstrFileName = FileHelper.GetFileName(ContentPath)
          lstrFileName = Path.GetFileName(ContentPath)
        End If

        lstrSerializationPath = Me.Document.SerializationPath
        'lstrVersionPath = String.Format("{0}\{1}", FileHelper.GetDirectoryName(lstrSerializationPath), Version.ID)
        lstrVersionPath = $"{Path.GetDirectoryName(lstrSerializationPath)}\{Version.ID}"

        lstrVersionPath = Helper.RemoveExtraBackSlashesFromFilePath(lstrVersionPath)

        ArchiveContentPath = String.Format("{0}\{1}", lstrVersionPath, lstrFileName)
        RelativePath = String.Format("{0}\{1}", Version.ID, lstrFileName)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Gets additional content properties available for this platform.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetExtendedContentProperties() As ECMProperties
      Try

        ' Add the MIMEType
        Dim lobjContentProperties As New ECMProperties From {
          PropertyFactory.Create(PropertyType.ecmString, "MIMEType", Me.MIMEType)
        }

        Return lobjContentProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function DetermineAvailability() As Boolean
      Try
        If mobjContentData IsNot Nothing AndAlso mobjContentData.Length AndAlso mobjContentData.CanRead Then
          Return True
        Else
          'Return New IO.FileStream(Me.ContentPath, FileMode.Create)
          If (Not String.IsNullOrEmpty(Me.ContentPath)) AndAlso (File.Exists(Me.ContentPath)) Then
            Return True
          Else
            Return False
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#End Region

#Region "ICloneable Implementation"

    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjContent As New Content()
      Try
        With lobjContent
          .ContentPath = Me.ContentPath.Clone
          .Data = Me.Data
          .Identifier = Me.Identifier.Clone
          .Metadata = Me.Metadata
          .MIMEType = Me.MIMEType.Clone
          .PreviousStorageType = Me.PreviousStorageType
          .RelativePath = Me.RelativePath.Clone
          .StorageType = Me.StorageType
        End With
        Return lobjContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace