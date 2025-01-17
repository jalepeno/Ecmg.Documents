'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Configuration.ConfigurationManager
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Ionic.Zip

#End Region

Namespace Core

  Partial Public Class Repository
    Implements ISerialize
    Implements IArchivable(Of Repository)
    Implements IXmlSerializable

#Region "Class Variables"

    Private mobjProvider As Providers.CProvider
    Private mobjProviderSystem As Providers.ProviderSystem
    Private mobjContentSource As Providers.ContentSource

    'Private mobjZipFile As ZipFile
    Private mobjChildStreams As List(Of ZipFileStream)

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpContentSource As Providers.ContentSource)
      Try

        If lpContentSource Is Nothing Then
          Throw New ArgumentNullException(NameOf(lpContentSource), "A valid value must be supplied for lpContentSource")
        End If

        mstrName = lpContentSource.Name
        mobjContentSource = lpContentSource

        If lpContentSource.Provider Is Nothing Then
          Throw New Exceptions.ProviderNotInitializedException(lpContentSource)
        End If

        If lpContentSource.Provider.IsInitialized = False Then
          lpContentSource.Provider.InitializeProvider(lpContentSource)
        End If

        If lpContentSource.Provider.IsConnected = False Then
          lpContentSource.Provider.Connect()
        End If

        InitializeProviderData(lpContentSource.Provider)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpContentSource As Providers.ContentSource, ByVal lpProvider As Providers.CProvider)
      Try

        If lpContentSource Is Nothing Then
          Throw New ArgumentNullException(NameOf(lpContentSource), "A valid value must be supplied for lpContentSource")
        End If

        mstrName = lpName
        mobjContentSource = lpContentSource
        InitializeProviderData(lpProvider)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpRepositoryStream As Stream)
      Try
        Dim lstrErrorMessage As String = ""
        Dim lobjRepository As Repository = Nothing
        lobjRepository = Deserialize(lpRepositoryStream, lstrErrorMessage)
        Helper.AssignObjectProperties(lobjRepository, Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFilePath As String)

      'If lstrErrorMessage.Length > 0 Then
      '  Throw New ArgumentException(String.Format("Unable to create repository object from xml file: {0}", lstrErrorMessage))
      'End If

      Try
        Dim lstrErrorMessage As String = ""
        Dim lobjRepository As Repository = Nothing
        lobjRepository = Deserialize(lpFilePath, lstrErrorMessage)
        Helper.AssignObjectProperties(lobjRepository, Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Object identifying the standard properties of a provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProviderSystem() As Providers.ProviderSystem
      Get
        Return mobjProviderSystem
      End Get
      Set(ByVal value As Providers.ProviderSystem)
        mobjProviderSystem = value
      End Set
    End Property

    ''' <summary>
    ''' The content source instance used when connecting to the repository
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ContentSource() As Providers.ContentSource
      Get
        Return mobjContentSource
      End Get
      Set(ByVal value As Providers.ContentSource)
        mobjContentSource = value
      End Set
    End Property

    Public ReadOnly Property ConnectionString As String
      Get
        Try
          If ContentSource IsNot Nothing Then
            Return ContentSource.ConnectionString
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property DefaultFileName As String
      Get
        Try
          Return Helper.CleanFile(String.Format("{0}.rfa", Me.Name), "_")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    'Public ReadOnly Property DefaultFilePath As String
    '  Get
    '    Try
    '      Return String.Format("{0}\{1}", FileHelper.Instance.RepositoriesPath, DefaultFileName)
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

    ''' <summary>
    ''' The provider used by the content source when connecting to the repository
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property Provider() As Providers.CProvider
      Set(ByVal value As Providers.CProvider)
        If value IsNot Nothing Then
          InitializeProviderData(value)
        End If
      End Set
    End Property

#End Region

#Region "Friend Properties"

    Friend ReadOnly Property ChildStreams As List(Of ZipFileStream)
      Get
        Return mobjChildStreams
      End Get
    End Property

#End Region

#Region "Public Methods"

    'Public Shared Sub ClearTempFiles(ByVal lpTargetDirectory As String)
    '  Try
    '    Helper.ClearTempFiles(lpTargetDirectory, "*.rif", "*.dcf", "*.cvl")
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Function GetAllChoiceLists() As ChoiceLists.ChoiceLists
      Try
        Dim lobjChoiceLists As ChoiceLists.ChoiceLists = Me.Properties.GetAllChoiceLists

        For Each lobjDocumentClass As DocumentClass In Me.DocumentClasses
          lobjChoiceLists.AddRange(lobjDocumentClass.GetAllChoiceLists)
        Next

        lobjChoiceLists.Sort()

        Return lobjChoiceLists

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Sub InitializeProviderData(ByVal lpProvider As Providers.CProvider)
      Try
        mobjProvider = lpProvider
        mobjProviderSystem = mobjProvider.ProviderSystem
        If lpProvider.SupportsInterface(ProviderClass.Classification) Then
          mobjDocumentClasses = CType(mobjProvider, Providers.IClassification).DocumentClasses
          mobjDocumentClasses?.Sort()
          mobjProperties = CType(mobjProvider, Providers.IClassification).ContentProperties
          ' Get all the subscribed classes
          mobjProperties.GetSubscribedClasses(mobjDocumentClasses)
        Else
          Throw New InvalidContentSourceException(
          String.Format("The abstract repository object can not be generated.  The supplied provider '{0}' does not implement the IClassification interface.", lpProvider.Name))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ISerialize Implementation"

    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DefaultFileExtension() As String Implements ISerialize.DefaultFileExtension
      Get
        Return REPOSITORY_INFORMATION_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpRepositoryStream As Stream, Optional ByRef lpErrorMessage As String = "") As Object
      Try
        ' Treat this is as zipped RIF
        Using lobjZipFile As ZipFile = ZipFile.Read(lpRepositoryStream)
          If lobjZipFile Is Nothing Then
            Throw New InvalidOperationException("Failed to read repository stream")
          End If
          If lobjZipFile.Entries.Count = 0 Then
            Throw New InvalidOperationException("No entries were found in the repository archive")
          End If
          If mobjRepositoryStream Is Nothing Then
            mobjRepositoryStream = New MemoryStream
          End If
          lobjZipFile.Item(0).Extract(mobjRepositoryStream)
          Return Serializer.Deserialize.FromZippedStream(mobjRepositoryStream, lobjZipFile, Me.GetType)
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try

        ' See if the request is using a zipped rif
        Dim lstrExtension As String = IO.Path.GetExtension(lpFilePath)

        Select Case lstrExtension.ToLower
          Case ".zrif", ".zrf", ".zip", ".rfa"
            ' Treat this is as zipped RIF
            Using lobjZipFile As New ZipFile(lpFilePath)
              'mobjZipFile = New ZipFile(lpFilePath)
              mobjRepositoryStream = New MemoryStream '(mobjZipFile.Entries(0).UncompressedSize)

              ' Set the source type, we will need this inside ReadXml
              mobjSerializationSourceType = SourceType.Stream

              Dim lstrRifFileName As String = IO.Path.GetFileName(lpFilePath).Replace(".rfa", ".rif")
              'mobjZipFile.Extract(lstrRifFileName, mobjRepositoryStream)
              'mobjZipFile.Item(lstrRifFileName).Extract(mobjRepositoryStream)
              lobjZipFile.Item(lstrRifFileName).Extract(mobjRepositoryStream)
              'Return Serializer.Deserialize.FromZippedStream(mobjRepositoryStream, mobjZipFile, Me.GetType)
              Return Serializer.Deserialize.FromZippedStream(mobjRepositoryStream, lobjZipFile, Me.GetType)
              'lobjZipFile.ExtractAll(IO.Path.GetDirectoryName(lpFilePath), True)
              'Return Serializer.Deserialize.XmlFile(lpFilePath.Replace(lstrExtension, ".rif"), Me.GetType)
            End Using

          Case Else
            Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
        End Select


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        'Return Nothing
        Throw
      End Try
    End Function

    Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Return Serializer.Serialize.Xml(Me)
    End Function

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Serializes a repository object to a file
    ''' </summary>
    ''' <param name="lpFilePath"></param>
    ''' <param name="lpZipped">Specifies whether or not to serialize to a zipped rif (zrif)</param>
    ''' <param name="lpDeleteOriginals">If lpZipped is true this specifies whether 
    ''' or not to delete the rif file and its component files 
    ''' (otherwise this value is ignored)</param>
    ''' <remarks></remarks>
    Public Sub Serialize(ByRef lpFilePath As String,
                         ByVal lpZipped As Boolean,
                         Optional ByVal lpDeleteOriginals As Boolean = True)
      Try

        If lpZipped = False Then
          RepositorySerializer.XmlFile(Me, lpFilePath)
        Else
          Archive(lpFilePath.Replace(REPOSITORY_INFORMATION_FILE_EXTENSION,
                                     REPOSITORY_INFORMATION_FILE_ARCHIVE_EXTENSION), lpDeleteOriginals)
        End If

        'Serializer.Serialize.XmlFile(Me, lpFilePath)

        'If lpZipped = True Then
        '  'ZipRif(lpFilePath, lpDeleteOriginals)

        'Else
        '  RepositorySerializer.XmlFile(Me, lpFilePath)
        'End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      ' Set the extension if necessary
      ' We want the default file extension on an ecmdocument to be .rif for
      ' "Repository Information File"

      Try

        ' If we were passed a rfa file extension then save it as a repository file archive
        If IO.Path.GetExtension(lpFilePath).EndsWith(REPOSITORY_INFORMATION_FILE_ARCHIVE_EXTENSION) Then
          'Serialize(lpFilePath, True)
          Archive(lpFileExtension, True)
          Exit Sub
        Else
          If lpFileExtension.Length = 0 Then
            ' No override was provided
            If lpFilePath.EndsWith(REPOSITORY_INFORMATION_FILE_EXTENSION) = False Then
              lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & REPOSITORY_INFORMATION_FILE_EXTENSION
            End If

          End If

          RepositorySerializer.XmlFile(Me, lpFilePath)

          ' Until we update CTStudio to give a choice we will create both the regular and the zipped rif
          Dim lblnCompressRifOutput As Boolean = False
          If AppSettings("CompressRifOutput") <> String.Empty Then
            lblnCompressRifOutput = CBool(AppSettings("CompressRifOutput"))
          End If
          If lblnCompressRifOutput = True Then
            ZipRif(lpFilePath, True)
          End If


        End If


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      If lpWriteProcessingInstruction = True Then
        RepositorySerializer.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
      Else
        RepositorySerializer.XmlFile(Me, lpFilePath)
      End If
    End Sub

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Return Serializer.Serialize.XmlString(Me)
    End Function

#Region "ZipRif Methods"

    ''' <summary>
    ''' Zips up the RIF file and all of its component files
    ''' </summary>
    ''' <param name="lpRifFilePath">The fully qualified path to the rif file</param>
    ''' <remarks></remarks>
    Public Shared Sub ZipRif(ByRef lpRifFilePath As String)
      Try
        ZipRif(lpRifFilePath, IO.Path.GetDirectoryName(lpRifFilePath), False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Zips up the RIF file and all of its component files to the specified output path
    ''' </summary>
    ''' <param name="lpRifFilePath">The fully qualified path to the rif file</param>
    ''' <param name="lpOutPutFilePath">The directory to save the new file to</param>
    ''' <remarks></remarks>
    Public Shared Sub ZipRif(ByRef lpRifFilePath As String,
                   ByVal lpOutPutFilePath As String)
      Try
        ZipRif(lpRifFilePath, lpOutPutFilePath, False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Zips up the RIF file and all of its component files
    ''' </summary>
    ''' <param name="lpRifFilePath">The fully qualified path to the rif file</param>
    ''' <param name="lpDeleteOriginals">Specifies whether or not to delete the original files</param>
    ''' <remarks></remarks>
    Public Shared Sub ZipRif(ByRef lpRifFilePath As String,
                   ByVal lpDeleteOriginals As Boolean)
      Try
        ZipRif(lpRifFilePath, IO.Path.GetDirectoryName(lpRifFilePath), lpDeleteOriginals)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Zips up the RIF file and all of its component files to the specified output path
    ''' </summary>
    ''' <param name="lpRifFilePath">The fully qualified path to the rif file</param>
    ''' <param name="lpOutPutFilePath">The directory to save the new file to</param>
    ''' <param name="lpDeleteOriginals">Specifies whether or not to delete the original files</param>
    ''' <remarks></remarks>
    Public Shared Sub ZipRif(ByRef lpRifFilePath As String,
                       ByVal lpOutPutFilePath As String,
                       ByVal lpDeleteOriginals As Boolean)
      Try
        Dim lobjZipFile As New ZipFile
        Dim lstrRifFolder As String = IO.Path.GetDirectoryName(lpRifFilePath)
        Dim lobjRifFolderInfo As New DirectoryInfo(lstrRifFolder)
        Dim lstrRepositoryName As String = IO.Path.GetFileNameWithoutExtension(lpRifFilePath)
        Dim lstrOutputFileName As String = String.Empty
        Dim lstrOutputFilePath As String = String.Empty
        Dim lstrDocClassFileFilter As String = String.Empty
        Dim lstrChoiceListFileFilter As String = String.Empty

        With lobjZipFile
          ' Add the RIF file
          .AddFile(lpRifFilePath, "")

          ' Add each document class file
          lstrDocClassFileFilter = String.Format("{0}*.{1}",
                                                 lstrRepositoryName,
                                                 DocumentClass.DOCUMENT_CLASS_FILE_EXTENSION)
          For Each lobjDocClassFile As FileInfo In lobjRifFolderInfo.GetFiles(lstrDocClassFileFilter)
            .AddFile(lobjDocClassFile.FullName, "")
          Next

          ' Add each ChoiceList file
          lstrChoiceListFileFilter = String.Format("{0}*.{1}", lstrRepositoryName, ChoiceLists.ChoiceList.CHOICE_LIST_FILE_EXTENSION)
          For Each lobjChoiceListFile As FileInfo In lobjRifFolderInfo.GetFiles(lstrChoiceListFileFilter)
            .AddFile(lobjChoiceListFile.FullName, "")
          Next

          ' Prepare the save location
          If IO.Directory.Exists(lpOutPutFilePath) = False Then
            IO.Directory.CreateDirectory(lpOutPutFilePath)
          End If

          lstrOutputFileName = String.Format("{0}.{1}", IO.Path.GetFileNameWithoutExtension(lpRifFilePath),
                                             REPOSITORY_INFORMATION_FILE_ARCHIVE_EXTENSION)
          lstrOutputFilePath = String.Format("{0}\{1}",
                                             lpOutPutFilePath, lstrOutputFileName)

          lstrOutputFilePath = Helper.RemoveExtraBackSlashesFromFilePath(lstrOutputFilePath)

          ' Save the zip file with a .zrif extension
          .Save(lstrOutputFilePath)


        End With

        ' Delete the original files if requested
        If lpDeleteOriginals = True Then

          ' Delete the Document Class (dcf) files
          For Each lobjDocClassFile As FileInfo In lobjRifFolderInfo.GetFiles(lstrDocClassFileFilter)
            IO.File.Delete(lobjDocClassFile.FullName)
          Next

          ' Delete the Choice List (cvl) files
          For Each lobjChoiceListFile As FileInfo In lobjRifFolderInfo.GetFiles(lstrChoiceListFileFilter)
            IO.File.Delete(lobjChoiceListFile.FullName)
          Next

          ' Delete the rif file
          IO.File.Delete(lpRifFilePath)

        End If

        ' Write the output file path back to the Rif file path
        lpRifFilePath = lstrOutputFilePath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "Private Classes"

    Private Class RepositorySerializer
      Inherits Serializer.Serialize

#Region "Overloaded Methods"

      Overloads Shared Sub XmlFile(ByVal lpRepository As Repository,
        ByVal lpFilePath As String,
        Optional ByVal lpSchemaLocation As String = "",
        Optional ByVal lpDeclaration() As String = Nothing,
        Optional ByVal lpWriteProcessingInstruction As Boolean = False,
        Optional ByVal lpXSLPath As String = "",
        Optional ByVal lpCleanChar As Char = "_")

        Try
          ' Serialize the file
          Serializer.Serialize.XmlFile(lpRepository, lpFilePath, lpSchemaLocation, lpDeclaration, lpWriteProcessingInstruction, lpXSLPath, lpCleanChar)

          ' The XML is unfortunately not formatted well
          ' We need to clean it up
          Dim lobjXMLDocument As New XmlDocument
          lobjXMLDocument.Load(lpFilePath)
          lobjXMLDocument = Helper.FormatXmlDocument(lobjXMLDocument)

          lobjXMLDocument.Save(lpFilePath)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Sub
#End Region
    End Class

#End Region

#End Region

#End Region

#Region "IArchivable(Of Repository) Implementation"

    Public ReadOnly Property DefaultArchiveFileExtension() As String Implements IArchivable(Of Repository).DefaultArchiveFileExtension
      Get
        Return REPOSITORY_INFORMATION_FILE_ARCHIVE_EXTENSION
      End Get
    End Property

    Public Function FromArchive(ByVal archivePath As String) As Repository Implements IArchivable(Of Repository).FromArchive
      Try

        'mobjZipFile = New ZipFile(archivePath)
        Using lobjZipFile As New ZipFile(archivePath)
          mobjRepositoryStream = New MemoryStream '(mobjZipFile.Entries(0).UncompressedSize)

          ' Set the source type, we will need this inside ReadXml
          mobjSerializationSourceType = SourceType.Stream

          Dim lstrRifFileName As String = IO.Path.GetFileName(archivePath).Replace(".rfa", ".rif")
          'mobjZipFile.Extract(lstrRifFileName, mobjRepositoryStream)
          'mobjZipFile.Item(lstrRifFileName).Extract(mobjRepositoryStream)
          lobjZipFile.Item(lstrRifFileName).Extract(mobjRepositoryStream)
          'Return Serializer.Deserialize.FromZippedStream(mobjRepositoryStream, mobjZipFile, Me.GetType)
          Return Serializer.Deserialize.FromZippedStream(mobjRepositoryStream, lobjZipFile, Me.GetType)
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function FromArchive(ByVal archivePath As String, ByVal password As String) As Repository Implements IArchivable(Of Repository).FromArchive
      Try
        ' Coming Soon
        Throw New NotImplementedException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function FromArchive(ByVal stream As System.IO.Stream) As Repository Implements IArchivable(Of Repository).FromArchive
      Try
        ' Coming Soon
        Throw New NotImplementedException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function FromArchive(ByVal stream As System.IO.Stream, ByVal password As String) As Repository Implements IArchivable(Of Repository).FromArchive
      Try
        Using lobjZipFile As ZipFile = ZipFile.Read(stream)

          If Not String.IsNullOrEmpty(password) Then
            lobjZipFile.Password = password
          End If
          Dim lobjRepositoryStream As New MemoryStream
          'mobjRepositoryStream = New MemoryStream
          ' Set the source type, we will need this inside ReadXml
          'mobjSerializationSourceType = SourceType.Stream
          'Dim lstrRifFileName As String = IO.Path.GetFileName(archivePath).Replace(".rfa", ".rif")
          Dim lstrFileNames = From lstrFileName In lobjZipFile.EntryFileNames Where lstrFileName.EndsWith(Repository.REPOSITORY_INFORMATION_FILE_EXTENSION)
          For Each lstrFileName As String In lstrFileNames
            lobjZipFile.Item(lstrFileName).Extract(lobjRepositoryStream)
            Exit For
          Next
          Return Serializer.Deserialize.FromZippedStream(lobjRepositoryStream, lobjZipFile, GetType(Repository))
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Archive(ByVal archivePath As String) Implements IArchivable(Of Repository).Archive
      Try

        Archive(archivePath, True, String.Empty)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Archive(ByVal archivePath As String, ByVal removeContainedFiles As Boolean) Implements IArchivable(Of Repository).Archive
      Try

        Archive(archivePath, removeContainedFiles, String.Empty)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Archive(ByVal archivePath As String, ByVal removeContainedFiles As Boolean, ByVal password As String) Implements IArchivable(Of Repository).Archive
      Try



        Dim lobjZipfileStream As IO.Stream = ToArchiveStream(password)
        Dim lobjZipFile As ZipFile = ZipFile.Read(lobjZipfileStream)

        If password IsNot Nothing AndAlso password.Length > 0 Then
          lobjZipFile.Password = password
        End If

        lobjZipFile.Save(archivePath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToArchiveStream() As System.IO.Stream Implements IArchivable(Of Repository).ToArchiveStream
      Try
        Return ToArchiveStream(String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToArchiveStream(ByVal password As String) As System.IO.Stream Implements IArchivable(Of Repository).ToArchiveStream
      Try

        Dim lobjRepositoryStream As IO.Stream
        Dim lobjOutputStream As New MemoryStream
        Dim lobjZipFile As New ZipFile
        Dim lstrRfa As String = String.Format("{0}.{1}", Name, Repository.REPOSITORY_INFORMATION_FILE_EXTENSION)

        lobjRepositoryStream = RepositorySerializer.ToStream(Me) ' .XmlFile(Me, lpFilePath)

        lobjZipFile.AddEntry(lstrRfa, lobjRepositoryStream)

        If ChildStreams IsNot Nothing Then
          For Each lobjChildStream As ZipFileStream In ChildStreams
            If lobjZipFile.ContainsEntry(lobjChildStream.FileName.TrimStart("\")) = False Then
              lobjZipFile.AddEntry(lobjChildStream.FileName, lobjChildStream.Stream)
            End If
          Next
          lobjZipFile.SortEntriesBeforeSaving = True
        End If

        If Not String.IsNullOrEmpty(password) Then
          lobjZipFile.Password = password
        End If

        lobjZipFile.Save(lobjOutputStream)

        If lobjOutputStream.CanSeek Then
          lobjOutputStream.Position = 0
        End If

        Return lobjOutputStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Sub ToPdf(pdfPath As String)
    '  Try
    '    Dim lobjRepositoryRendition As New RepositoryRendition(Me)
    '    lobjRepositoryRendition.ToPdf(pdfPath)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Public Function ToPdfStream() As Stream
    '  Try
    '    Dim lobjRepositoryRendition As New RepositoryRendition(Me)
    '    Return lobjRepositoryRendition.ToPdfStream
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Sub ToRtf(pdfPath As String)
    '  Try
    '    Dim lobjRepositoryRendition As New RepositoryRendition(Me)
    '    lobjRepositoryRendition.ToRtf(pdfPath)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Public Function ToRtfStream() As Stream
    '  Try
    '    Dim lobjRepositoryRendition As New RepositoryRendition(Me)
    '    Return lobjRepositoryRendition.ToRtfStream
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Function ToByteArray() As Byte()
      Try
        Dim lobjReturnArray As Byte()
        Using lobjRepositoryStream As Stream = ToArchiveStream()
          lobjReturnArray = Helper.ReadStreamToByteArray(lobjRepositoryStream, lobjRepositoryStream.Length)
        End Using
        Return lobjReturnArray
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      ' As per the Microsoft guidelines this is not implemented
      Return Nothing
    End Function

    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Try

        Throw New NotImplementedException()

        'If reader.GetType.Name <> "EnhancedStreamReader" AndAlso reader.GetType.IsSubclassOf(GetType(EnhancedStreamReader)) Then
        '  ApplicationLogging.WriteLogEntry("Unable to provide enhanced deserialization for object")
        'End If

        'Dim lstrFileName As String = String.Empty
        'Dim lstrSerializationFolderPath As String = String.Empty

        'If Helper.CallStackContainsMethodName("FromStream", "FromZippedStream") Then
        '  mobjSerializationSourceType = SourceType.Stream
        'End If

        'If mobjSerializationSourceType = SourceType.File Then
        '  ' Get the file name from the Base URI
        '  lstrFileName = reader.BaseURI.Replace("file:///", "").Replace("/", "\")
        '  ' Get the file folder
        '  lstrSerializationFolderPath = Path.GetDirectoryName(lstrFileName)
        'End If

        'Dim lstrDocumentClassFileName As String = String.Empty
        'Dim lstrDocumentClassFile As String = String.Empty
        'Dim lobjProviderSystem As ProviderSystem = Nothing
        'Dim lobjContentSource As ContentSource = Nothing
        'Dim lobjDocumentClass As DocumentClass = Nothing
        'Dim lobjClassificationProperty As ClassificationProperty = Nothing
        'Dim lstrElementString As String = String.Empty
        'Dim lobjPropertyNameNode As XmlNode = Nothing
        'Dim lstrPropertyName As String = String.Empty
        'Dim lobjXmlDocument As New XmlDocument
        'Dim lobjTargetStream As IO.Stream = Nothing
        'Dim lobjZipFile As ZipFile = Nothing
        'Dim lobjDocClassEntry As ZipEntry = Nothing

        'lobjXmlDocument.Load(reader)

        'With lobjXmlDocument
        '  ' Get the name
        '  Me.Name = .DocumentElement.GetAttribute("Name")

        '  ' For backwards compatibility
        '  If Me.Name = String.Empty Then
        '    Me.Name = .SelectSingleNode("Repository/Name").InnerText
        '  End If

        '  ' Read the ProviderSystem
        '  ' Construct the provider system
        '  lobjProviderSystem = Serializer.Deserialize.SoapString(
        '    .SelectSingleNode("//ProviderSystem").OuterXml, GetType(ProviderSystem))
        '  ' Add the provider system
        '  Me.ProviderSystem = lobjProviderSystem

        '  ' Read the ContentSource

        '  ' Read the ContentSource element
        '  Dim lobjContentSourceNode As XmlNode = .SelectSingleNode("//ContentSource")

        '  If lobjContentSourceNode IsNot Nothing Then
        '    lstrElementString = lobjContentSourceNode.OuterXml

        '    ' Remove any reference to xsd if present
        '    lstrElementString = CleanUpXsd(lstrElementString)

        '    ' Construct the ContentSource
        '    lobjContentSource = Serializer.Deserialize.SoapString(lstrElementString,
        '                                                          GetType(ContentSource))
        '    ' Add the ContentSource
        '    Me.ContentSource = lobjContentSource
        '  Else
        '    ApplicationLogging.WriteLogEntry("Failed to read content source from repository file.", TraceEventType.Warning, 61392)
        '  End If

        '  ' Read the DocumentClasses
        '  Dim lobjDocumentClassNodes As XmlNodeList = .SelectNodes("//DocumentClass")
        '  For Each lobjDocumentClassNode As XmlNode In lobjDocumentClassNodes

        '    Try
        '      ' Get the name of the document class file
        '      lstrDocumentClassFileName = CType(lobjDocumentClassNode, XmlElement).GetAttribute("href")

        '      ' For backwards compatibility
        '      If lstrDocumentClassFileName = String.Empty Then
        '        ' Construct a new document class file using the xml file path
        '        lobjDocumentClass = Serializer.Deserialize.SoapString(CleanUpXsd(lobjDocumentClassNode.OuterXml),
        '                                                              GetType(DocumentClass))
        '        If lobjDocumentClass.Name = String.Empty Then
        '          ' We changed this to an attribute at one point
        '          Dim lobjNameNode As XmlNode = lobjDocumentClassNode.SelectSingleNode("Name")
        '          If lobjNameNode IsNot Nothing Then
        '            lobjDocumentClass.Name = lobjNameNode.InnerText
        '          End If
        '          'Dim lobjNameAttr As XmlAttribute = lobjDocumentClassNode.Attributes("name")
        '          'If (lobjNameAttr IsNot Nothing) Then
        '          '  lobjDocumentClass.Name = lobjNameAttr.InnerText
        '          'End If
        '        End If
        '      Else
        '        If mobjSerializationSourceType = SourceType.File Then
        '          ' Build the fully qualified path
        '          lstrDocumentClassFile = String.Format("{0}{1}", lstrSerializationFolderPath,
        '                                                lstrDocumentClassFileName)

        '          ' Make sure we have a clean file name
        '          lstrDocumentClassFile = Helper.CleanFile(lstrDocumentClassFile, "-")

        '          ' Construct a new document class file using the xml file path
        '          lobjDocumentClass = New DocumentClass(lstrDocumentClassFile)
        '        Else
        '          ' Get the document class file out of the zip file
        '          If TypeOf (reader) Is XmlZipStreamReader Then
        '            lobjZipFile = CType(reader, XmlZipStreamReader).ZipFile
        '            lobjTargetStream = New MemoryStream

        '            If lobjZipFile.ContainsEntry(lstrDocumentClassFileName) = False Then
        '              ' Make sure we have a clean file name
        '              lstrDocumentClassFileName = Helper.CleanFile(lstrDocumentClassFileName, "-").Replace("\", String.Empty)
        '            End If

        '            ' Try one more time
        '            If lobjZipFile.ContainsEntry(lstrDocumentClassFileName) = False Then
        '              Dim lstrDocumentClassName As String = CType(lobjDocumentClassNode, XmlElement).GetAttribute("name")
        '              Throw New Exceptions.DocumentClassNotInitializedException(
        '                String.Format("Unable to extract document class file for class '{0}' from rfa file.",
        '                              lstrDocumentClassName), lstrDocumentClassName, Nothing)
        '            End If

        '            lobjDocClassEntry = lobjZipFile.Item(lstrDocumentClassFileName)
        '            lobjDocClassEntry.Extract(lobjTargetStream)
        '            lobjDocumentClass = New DocumentClass(lobjTargetStream, lobjZipFile)

        '          End If
        '        End If
        '      End If
        '      ' Add the document class to the collection
        '      Me.DocumentClasses.Add(lobjDocumentClass)
        '    Catch ex As Exception
        '      Throw New DocumentClassNotInitializedException("Unable to initialize document class from file.",
        '                                                     lstrDocumentClassFileName, lobjContentSource, ex)
        '    End Try

        '  Next

        '  ' Read the Properties
        '  'Dim lobjPropertyNodes As XmlNodeList = .SelectNodes("//Properties//ClassificationProperty")
        '  Dim lobjPropertyNodes As XmlNodeList = .SelectNodes("Repository/Properties/*")
        '  For Each lobjPropertyNode As XmlNode In lobjPropertyNodes
        '    Try
        '      ' Read the ClassificationProperty element
        '      lstrElementString = lobjPropertyNode.OuterXml

        '      ' Get the property name, we may need it in case of an exception
        '      lobjPropertyNameNode = lobjPropertyNode.SelectSingleNode("Name")
        '      If lobjPropertyNameNode IsNot Nothing Then
        '        lstrPropertyName = lobjPropertyNameNode.InnerText
        '      Else
        '        ' If we can't get the node then just default to the entire element string
        '        lstrPropertyName = lstrElementString
        '      End If

        '      ' Remove any reference to xsd if present
        '      lstrElementString = CleanUpXsd(lstrElementString)
        '      ' Construct the ClassificationProperty
        '      'lobjClassificationProperty = Serializer.Deserialize.SoapString( _
        '      '  lstrElementString, GetType(ClassificationProperty))
        '      lobjClassificationProperty = ClassificationProperty.CreateFromXmlelement(CType(lobjPropertyNode, XmlElement))

        '      ' Add the property to the collection
        '      Me.Properties.Add(lobjClassificationProperty)
        '    Catch ex As Exception
        '      Throw New PropertyNotInitializedException(String.Format("Unable to initialize property '{0}' from repository file.",
        '                  lstrPropertyName), lstrPropertyName, lobjDocumentClass, lobjContentSource, ex)
        '    End Try
        '  Next

        'End With

      Catch ex As Exception
        Dim lobjRepositoryException As RepositoryDeserializationException
        Select Case ex.GetType.Name
          Case "DocumentClassNotInitializedException"
            Dim lobjDocClassEx As DocumentClassNotInitializedException = CType(ex, DocumentClassNotInitializedException)
            lobjRepositoryException = New RepositoryDeserializationException(
              String.Format("Unable to deserialize repository '{0}' from file '{1}', there was a problem with the document class '{2}'.",
                            Me.Name, reader.BaseURI, lobjDocClassEx.DocumentClassName),
                            reader.BaseURI, ex)
          Case "PropertyNotInitializedException"
            Dim lobjPropertyEx As PropertyNotInitializedException = CType(ex, PropertyNotInitializedException)
            lobjRepositoryException = New RepositoryDeserializationException(
              String.Format("Unable to deserialize repository '{0}' from file '{1}', there was a problem with the property '{2}' in document class '{3}'.",
                            Me.Name, reader.BaseURI, lobjPropertyEx.Property, lobjPropertyEx.DocumentClass.Name),
                            reader.BaseURI, ex)
          Case Else
            lobjRepositoryException = New RepositoryDeserializationException(ex.Message, reader.BaseURI, ex)
        End Select

        ApplicationLogging.LogException(lobjRepositoryException, Reflection.MethodBase.GetCurrentMethod)
        '  Throw the exception to the caller
        Throw lobjRepositoryException
      End Try

    End Sub

    ''' <summary>
    ''' Writes the xml behind the repository to an XML file
    ''' </summary>
    ''' <param name="writer"></param>
    ''' <remarks></remarks>
    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        If writer.GetType.Name <> "EnhancedXmlTextWriter" AndAlso writer.GetType.IsSubclassOf(GetType(EnhancedXmlTextWriter)) Then
          ApplicationLogging.WriteLogEntry("Unable to provide enhanced serialization for object")
        End If

        Dim lstrRepositoryFileName As String

        If (TypeOf (writer) Is EnhancedXmlTextWriter = False) AndAlso (Helper.CallStackContainsMethodName("ToStream")) Then
          lstrRepositoryFileName = String.Format("{0}.{1}", Me.Name, REPOSITORY_INFORMATION_FILE_EXTENSION)
        Else
          lstrRepositoryFileName = CType(writer, EnhancedXmlTextWriter).FileName
        End If

        Dim lstrFileLocation As String = IO.Path.GetDirectoryName(lstrRepositoryFileName)
        Dim lstrDocumentClassFileName As String = String.Empty

        mobjChildStreams = New List(Of ZipFileStream)

        With writer

          ' Write a human readable file
          '.Formatting = Formatting.Indented

          ' Write the Repository Name element
          .WriteAttributeString("Name", Me.Name)

          '.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
          .WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")
          '.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
          .WriteAttributeString("xmlns", "xsd", Nothing, "http://www.w3.org/2001/XMLSchema")

          ' Write the ProviderSystem element
          If Me.ProviderSystem IsNot Nothing Then
            .WriteRaw(Me.ProviderSystem.ToXmlString)
          Else
            .WriteRaw((New ProviderSystem).ToXmlString)
          End If


          ' Write the ContentSource element
          If Me.ContentSource IsNot Nothing Then
            .WriteRaw(Me.ContentSource.ToXmlString)
          Else
            .WriteRaw((New ContentSource).ToXmlString)
          End If


          ' Open the DocumentClasses Element
          .WriteStartElement("DocumentClasses")

          ' Break them out into separate files
          For Each lobjDocumentClass As DocumentClass In Me.DocumentClasses

            ' Make sure the Repository is set
            lobjDocumentClass.Repository = Me

            ' Open the DocumentClasses Element
            .WriteStartElement("DocumentClass")

            ' Write the name of the document class as an attribute
            .WriteAttributeString("name", lobjDocumentClass.Name)

            ' Write the location of the documentclass file
            lstrDocumentClassFileName = String.Format("{0}_{1}.{2}",
                                                      Me.Name,
                                                      lobjDocumentClass.Name,
                                                      DocumentClass.DOCUMENT_CLASS_FILE_EXTENSION)

            ' It is possible that there were characters used in the 
            ' document class name that are not valid for a file name.
            ' Scrub any invalid file name characters from the file name
            lstrDocumentClassFileName = Helper.CleanFile(lstrDocumentClassFileName, "-", False)

            .WriteAttributeString("href", lstrDocumentClassFileName)

            If ChildStreams IsNot Nothing Then
              ChildStreams.Add(New ZipFileStream(lstrDocumentClassFileName, String.Empty, lobjDocumentClass.SerializeToStream))
            Else
              ' Serialize the document class
              lstrDocumentClassFileName = String.Format("{0}\{1}", lstrFileLocation, lstrDocumentClassFileName)
              lobjDocumentClass.Serialize(lstrDocumentClassFileName)
            End If

            'mobjChildStreams.Add(lobjDocumentClass.SerializeToStream)

            ' Close the DocumentClass element
            .WriteEndElement()
          Next

          ' Close the DocumentClasses element
          .WriteEndElement()

          ' Open the Properties Element
          .WriteStartElement("Properties")

          For Each lobjClassificationProperty As ClassificationProperty In Me.Properties

            ' Write the ClassificationProperty element
            .WriteRaw(lobjClassificationProperty.ToXmlString)

          Next
          ' Close the Properties element
          .WriteEndElement()
          '.Close()

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Takes a soap string and removes xsd:string attributes
    ''' </summary>
    ''' <param name="lpSoap">Xml soap string</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CleanUpXsd(ByVal lpSoap As String) As String
      Try
        ' Remove any reference to xsd if present
        Dim lstrSoap As String = lpSoap

        lstrSoap = lstrSoap.Replace(String.Format(" xsi:type={0}xsd:string{0}",
                                                                    ControlChars.Quote), "")

        lstrSoap = lstrSoap.Replace(String.Format(" xsi:type={0}xsd:int{0}",
                                                                    ControlChars.Quote), "")

        Return lstrSoap

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace