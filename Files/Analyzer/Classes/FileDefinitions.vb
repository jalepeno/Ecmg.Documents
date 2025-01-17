'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  FileDefinitions.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 2:13:11 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Reflection
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Ionic.Zip
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization

#End Region

Namespace Files
  <Serializable()> Public Class FileDefinitions
    Inherits List(Of FileDefinition)
    Implements IFileDefinitions

#Region "Class Variables"

    <NonSerialized()> Private mobjEnumerator As IEnumeratorConverter(Of IFileDefinition)
    Private mintFrontBlockSize As Integer

#End Region

#Region "IFileDefinitions Implementation"

    Public Overloads Sub Add(item As FileDefinition)
      Try
        If Not ContainsRawFileDefinition(item) Then 'MyBase.Contains(lobjDefinition) Then
          MyBase.Add(item)
          If Me.FrontBlockSize < item.FrontBlockSize Then
            mintFrontBlockSize = item.FrontBlockSize
          End If
        Else
          ApplicationLogging.WriteLogEntry(String.Format("{0} already exists.", item.DebuggerIdentifier()), MethodBase.GetCurrentMethod(), TraceEventType.Warning, 62361)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(item As IFileDefinition) Implements ICollection(Of IFileDefinition).Add
      Try
        If TypeOf (item) Is FileDefinition Then
          'Dim lobjDefinition As FileDefinition = item
          ''lobjDefinition.ItemsFound += 1
          'If Not ContainsRawFileDefinition(item) Then 'MyBase.Contains(lobjDefinition) Then
          '  MyBase.Add(lobjDefinition)
          '  If Me.FrontBlockSize < lobjDefinition.FrontBlockSize Then
          '    mintFrontBlockSize = lobjDefinition.FrontBlockSize
          '  End If
          'End If
          Add(DirectCast(item, FileDefinition))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements ICollection(Of IFileDefinition).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IFileDefinition) As Boolean Implements ICollection(Of IFileDefinition).Contains
      Try
        Return Contains(CType(item, FileDefinition))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Function ContainsRawFileDefinition(item As IFileDefinition, Optional ByRef lpFoundDefinition As IFileDefinition = Nothing) As Boolean
      Try
        Dim lobjDefinitions As List(Of FileDefinition) = Me
        'Dim lobjFileDefinition = lobjDefinitions.FirstOrDefault(Function(e) String.Compare(e.FileInfo.Extension, lpExtension, True) = 0)

        'If lobjFileDefinition IsNot Nothing Then
        '  If Not String.IsNullOrEmpty(lobjFileDefinition.FileInfo.MimeType) Then
        '    Return lobjFileDefinition.FileInfo.MimeType
        '  Else
        '    Return lpDefaultValue
        '  End If
        'Else
        '  Return lpDefaultValue
        'End If

        Dim lobjFileDefinition As IFileDefinition
        lobjFileDefinition = ItemByInfoHash(item.InfoHash)
        If lobjFileDefinition Is Nothing Then
          Return False
        Else
          lpFoundDefinition = lobjFileDefinition
          Return True
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IFileDefinition, arrayIndex As Integer) Implements ICollection(Of IFileDefinition).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads ReadOnly Property Count As Integer Implements ICollection(Of IFileDefinition).Count
      Get
        Try
          Return MyBase.Count
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of IFileDefinition).IsReadOnly
      Get
        Try
          Return False
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads Function Remove(item As IFileDefinition) As Boolean Implements ICollection(Of IFileDefinition).Remove
      Try
        Return MyBase.Remove(CType(item, FileDefinition))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As IEnumerator(Of IFileDefinition) Implements IEnumerable(Of IFileDefinition).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function IndexOf(item As IFileDefinition) As Integer Implements IList(Of IFileDefinition).IndexOf
      Try
        Return MyBase.IndexOf(CType(item, FileDefinition))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub Insert(index As Integer, item As IFileDefinition) Implements IList(Of IFileDefinition).Insert
      Try
        MyBase.Insert(index, CType(item, FileDefinition))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Default Public Overloads Property Item(index As Integer) As IFileDefinition Implements IList(Of IFileDefinition).Item
      Get
        Try
          Return MyBase.Item(index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IFileDefinition)
        Try
          MyBase.Item(index) = CType(value, FileDefinition)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads Sub RemoveAt(index As Integer) Implements IList(Of IFileDefinition).RemoveAt
      Try
        MyBase.RemoveAt(index)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overloads Sub Sort()
      Try
        MyBase.Sort()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub CreateFromLegacyXmlFiles(lpFolderPath As String)
      Try
        Clear()
        Dim lstrDefinitionFiles As String() = Directory.GetFiles(lpFolderPath, "*.trid.xml")
        Dim lobjFileDefinition As FileDefinition

        For Each lstrDefinitionFile As String In lstrDefinitionFiles
          lobjFileDefinition = FileDefinition.CreateFromLegacyXmlFile(lstrDefinitionFile)
          lobjFileDefinition.FileInfo.Popularity = Rating.C
          Add(lobjFileDefinition)
        Next

        Sort()

        Dim lobjPopularityReferenceSet As FileDefinitions = CreatePopularityReferenceSet("C:\Users\Russ\Documents\FileDefs11_Rated.txt")


        Dim lobjPopularityDefinition As FileDefinition
        ' Add the popularity rating to the master set
        For Each lobjFileDefinition In Me
          lobjPopularityDefinition = lobjPopularityReferenceSet.ItemByInfoHash(lobjFileDefinition.InfoHash)
          If lobjPopularityDefinition IsNot Nothing Then
            lobjFileDefinition.FileInfo.Popularity = lobjPopularityDefinition.FileInfo.Popularity
          Else
            'Beep()
          End If
          If lobjFileDefinition.FileInfo.RefUrl = "www.ecmg.com" Then
            lobjFileDefinition.FileInfo.Popularity = Rating.B
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function CreatePopularityReferenceSet(lpFilePath As String) As IFileDefinitions
      Try
        Dim lobjDefinitions As New FileDefinitions
        Dim lobjDefinition As FileDefinition
        Dim lstrLineParts() As String
        'Dim lstrExtension As String
        'Dim lstrFileType As String
        'Dim lstrMimeType As String
        'Dim lstrRefUrl As String
        'Dim lstrRemarks As String
        'Dim lstrPopularity As String
        'Dim lenuPopularity As Rating

        Dim lstrLines As List(Of String) = File.ReadAllLines(lpFilePath).ToList()

        For Each lstrLine As String In lstrLines
          lstrLineParts = lstrLine.Split(ControlChars.Tab)
          'lstrExtension = lstrLineParts(0)
          'lstrFileType = lstrLineParts(1)
          'lstrMimeType = lstrLineParts(2)
          'lstrRefUrl = lstrLineParts(3)
          'lstrRemarks = lstrLineParts(4)
          'lstrRemarks = lstrLineParts(5)
          'lstrPopularity = lstrLineParts(6)
          'lenuPopularity = [Enum].Parse(GetType(Rating), lstrPopularity)
          lobjDefinition = New FileDefinition()

          With lobjDefinition.FileInfo
            .Extension = lstrLineParts(0)
            .FileType = lstrLineParts(1)
            .MimeType = lstrLineParts(2)
            .RefUrl = lstrLineParts(3)
            .Remarks = lstrLineParts(4)
            .Popularity = [Enum].Parse(GetType(Rating), lstrLineParts(5).ToUpper())
          End With

          lobjDefinitions.Add(lobjDefinition)
        Next

        Return lobjDefinitions

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetMimeType(lpExtension As String, lpDefaultValue As String) As String Implements IFileDefinitions.GetMimeType
      Try

        Dim lobjDefinitions As List(Of FileDefinition) = Me
        Dim lobjFileDefinition As FileDefinition = lobjDefinitions.FirstOrDefault(Function(e) String.Compare(e.FileInfo.Extension, lpExtension, True) = 0)

        If lobjFileDefinition IsNot Nothing Then
          If Not String.IsNullOrEmpty(lobjFileDefinition.FileInfo.MimeType) Then
            Return lobjFileDefinition.FileInfo.MimeType
          Else
            Return lpDefaultValue
          End If
        Else
          Return lpDefaultValue
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ItemByInfoHash(lpInfoHash As String) As IFileDefinition
      Try
        Dim lobjDefinitions As List(Of FileDefinition) = Me
        Dim lobjFileDefinition As FileDefinition = lobjDefinitions.FirstOrDefault(Function(e) String.Compare(e.InfoHash, lpInfoHash) = 0)

        Return lobjFileDefinition

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ItemsByPopularity(lpPopularity As Rating) As IFileDefinitions
      Try
        Dim lobjReturnSet As New FileDefinitions
        Dim lobjResults As Object = From fileDef In Me Where fileDef.FileInfo.Popularity = lpPopularity Select fileDef

        For Each lobjFileDef As IFileDefinition In lobjResults
          lobjReturnSet.Add(lobjFileDef)
        Next

        Return lobjReturnSet

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub SaveToJson(lpFilePath As String, Optional lpIndented As Boolean = False)
      Try
        Dim lobjJsonSerializer As New JsonSerializer

        With lobjJsonSerializer
          .TypeNameHandling = TypeNameHandling.None
          .ContractResolver = New CamelCasePropertyNamesContractResolver()
          .StringEscapeHandling = StringEscapeHandling.EscapeNonAscii
          If lpIndented Then
            .Formatting = Xml.Formatting.Indented
          End If
        End With

        Using lobjStreamWriter As New StreamWriter(lpFilePath)
          Using lobjJsonWriter As New JsonTextWriter(lobjStreamWriter)
            lobjJsonSerializer.Serialize(lobjJsonWriter, Me)
          End Using
        End Using
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToJson(lpIndented As Boolean) As String
      Try

        Dim lobjJsonSettings As New JsonSerializerSettings()
        With lobjJsonSettings
          .TypeNameHandling = TypeNameHandling.None
          .ContractResolver = New CamelCasePropertyNamesContractResolver()
        End With

        'If Helper.IsRunningInstalled Then
        '  Return JsonConvert.SerializeObject(Me, Formatting.None, lobjJsonSettings)
        'Else
        '  Return JsonConvert.SerializeObject(Me, Formatting.Indented, lobjJsonSettings)
        'End If

        If lpIndented Then
          Return JsonConvert.SerializeObject(Me, Formatting.Indented, lobjJsonSettings)
        Else
          Return JsonConvert.SerializeObject(Me, Formatting.None, lobjJsonSettings)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Open(lpFilePath As String) As FileDefinitions
      Try
        ' Check to see if we have a valid path
        Helper.VerifyFilePath(lpFilePath, "lpFilePath", True)

        If ZipFile.IsZipFile(lpFilePath) = False Then
          Throw New Exceptions.InvalidPathException(
            String.Format("The file '{0}' is not a valid archive file", lpFilePath), lpFilePath)
        End If

        Dim lobjZipStream As MemoryStream = Serializer.Deserialize.ReadFileToMemory(lpFilePath)

        Return FromArchive(lobjZipStream)

        'Dim lobjFileDefinitions As FileDefinitions

        'Using lobjFileStream As New FileStream(lpFilePath, FileMode.Open)
        '  ' Create a binary formatter for this stream
        '  Dim lobjFormatter As New BinaryFormatter
        '  ' Deserialize the contents of the file stream.
        '  lobjFileDefinitions = DirectCast(lobjFormatter.Deserialize(lobjFileStream), FileDefinitions)
        'End Using

        'Return lobjFileDefinitions

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function FromArchive(ByVal stream As System.IO.Stream) As FileDefinitions
      Try

        stream.Position = 0
        Dim lobjZipFile As ZipFile = ZipFile.Read(stream)

        Return FromZip(lobjZipFile, String.Empty)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function FromZip(ByVal zip As Object,
                         ByVal password As String) As FileDefinitions
      Try
        If zip Is Nothing Then
          Throw New ArgumentNullException(NameOf(zip))
        ElseIf TypeOf (zip) Is ZipFile = False Then
          Throw New ArgumentException(
            String.Format("The type '{0}' was not expected, zip must be of type ZipFile.",
                          zip.GetType.Name), NameOf(zip))
        End If

        Dim lobjZipFile As ZipFile = zip

        If lobjZipFile IsNot Nothing Then

          ' If supplied, set the password
          If password IsNot Nothing AndAlso password.Length > 0 Then
            lobjZipFile.Password = password
          End If

          If lobjZipFile.Entries.Count = 0 OrElse
    lobjZipFile.Item(0).FileName <> "FileDefinitions.json" Then
            Throw New Exceptions.DocumentException(String.Empty,
                          String.Format("The archive '{0}' does not begin with a json Document.", lobjZipFile.Name))
          End If

          Dim lobjDocumentStream As New MemoryStream
          lobjZipFile.Item(0).Extract(lobjDocumentStream)

          Dim lobjFileDefinitions As FileDefinitions

          lobjFileDefinitions = FileDefinitions.FromJson(Helper.CopyStreamToString(lobjDocumentStream))

          Return lobjFileDefinitions

        Else
          Return Nothing
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Save(lpFilePath As String)
      Try

        ' Verify the requested path
        Helper.VerifyFilePath(lpFilePath, False)

        Using lobjArchiveMemoryStream As MemoryStream = ToArchiveStream()
          Using lobjFileStream As New FileStream(lpFilePath, FileMode.Create)
            lobjArchiveMemoryStream.WriteTo(lobjFileStream)

            lobjFileStream.Close()

          End Using
        End Using


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function FromJsonFile(lpFilePath As String) As FileDefinitions
      Try
        Dim lstrJson As String = Helper.ReadAllTextFromFile(lpFilePath)
        Return FromJson(lstrJson)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function FromJson(lpJsonString As String) As FileDefinitions
      Try
        Dim lobjJsonSettings As New JsonSerializerSettings()
        With lobjJsonSettings
          .TypeNameHandling = TypeNameHandling.None
          .ContractResolver = New CamelCasePropertyNamesContractResolver()
          .StringEscapeHandling = StringEscapeHandling.EscapeNonAscii
        End With

        Return JsonConvert.DeserializeObject(lpJsonString, GetType(FileDefinitions), lobjJsonSettings)

        'Return JsonConvert.DeserializeObject(lpJsonString, GetType(FileDefinitions), New FileDefinitionsConverter())

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToArchiveStream() As Stream
      Try
        Dim lobjOutputStream As New MemoryStream
        Dim lobjZipFile As New ZipFile

        With lobjZipFile
          .AddEntry("FileDefinitions.json", Me.ToJsonStream())
        End With

        lobjZipFile.Save(lobjOutputStream)

        If lobjOutputStream.CanSeek Then
          lobjOutputStream.Position = 0
        End If

        lobjZipFile.Dispose()

        lobjZipFile = Nothing

        Return lobjOutputStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Public Function ToJsonStream() As IO.Stream
      Try

        Return Helper.CopyStringToStream(Me.ToJson(False))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Friend Properties"

    Friend ReadOnly Property FrontBlockSize As Integer
      Get
        Try
          Return mintFrontBlockSize
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IFileDefinition)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IFileDefinition)(Me.ToArray, GetType(FileDefinition), GetType(IFileDefinition))
          End If
          Return mobjEnumerator
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

  End Class

End Namespace