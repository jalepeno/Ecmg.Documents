'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  FileDefinition.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 2:06:16 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Text
Imports System.Xml
Imports Documents.Utilities
Imports Microsoft.VisualBasic.CompilerServices
Imports Newtonsoft.Json

#End Region

Namespace Files

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  <Serializable()> Public Class FileDefinition
    Implements IFileDefinition
    Implements IComparable

#Region "Class Variables"

    Private mobjFileInfo As New FileInformation
    Private mobjHeaderPatterns As New HeaderPatterns
    Private mobjGlobalStrings As New List(Of ByteString)
    Private mintFrontBlockSize As Integer
    Private mlngItemsFound As Long

#End Region

#Region "Public Properties"

    <JsonProperty("fi")> Public Property FileInfo As IFileInformation Implements IFileDefinition.FileInfo
      Get
        Try
          Return mobjFileInfo
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IFileInformation)
        Try
          mobjFileInfo = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <JsonProperty("ps")> Public Property Patterns As IHeaderPatterns Implements IFileDefinition.Patterns
      Get
        Try
          Return mobjHeaderPatterns
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IHeaderPatterns)
        Try
          mobjHeaderPatterns = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <JsonProperty("gs")> Public Property GlobalStrings As IList(Of ByteString) Implements IFileDefinition.GlobalStrings
      Get
        Try
          Return mobjGlobalStrings
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IList(Of ByteString))
        Try
          mobjGlobalStrings = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Friend Property FrontBlockSize As Integer Implements IFileDefinition.FrontBlockSize
      Get
        Try
          Return mintFrontBlockSize
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintFrontBlockSize = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <JsonIgnore()>
    Public ReadOnly Property InfoHash As String Implements IFileDefinition.InfoHash
      Get
        Try
          Return mobjFileInfo.BasicHash
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <JsonIgnore()>
    Public Property ItemsFound As Long Implements IFileDefinition.ItemsFound
      Get
        Try
          Return mlngItemsFound
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Long)
        Try
          mlngItemsFound = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpFileInfo As IFileInformation, lpHeaderPatterns As IHeaderPatterns, lpGlobalStrings As ByteStrings)
      Try
        FileInfo = lpFileInfo
        Patterns = lpHeaderPatterns
        GlobalStrings = lpGlobalStrings
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try
        'Dim lintFindCountComparison As Integer = ItemsFound.CompareTo(obj.ItemsFound)

        '' First compare by the number of items found
        'If lintFindCountComparison <> 0 Then
        '  Return lintFindCountComparison
        'Else
        '  ' Then compare by the actual file information
        '  Return mobjFileInfo.CompareTo(obj.fileinfo)
        'End If

        Return mobjFileInfo.CompareTo(obj.fileinfo)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Function CompareRawFileDefinition(obj As IFileDefinition) As Integer
      Try
        Dim lintFileInfoComparison As Integer = mobjFileInfo.CompareTo(obj.FileInfo)

        If lintFileInfoComparison <> 0 Then
          Return lintFileInfoComparison
        End If

        Dim lintPatternsComparison As Integer = mobjHeaderPatterns.CompareTo(obj.Patterns)

        If lintPatternsComparison <> 0 Then
          Return lintPatternsComparison
        End If

        Dim lintByteStringCountComparison As Integer = GlobalStrings.Count.CompareTo(obj.GlobalStrings.Count)

        If lintByteStringCountComparison <> 0 Then
          Return lintByteStringCountComparison
        End If

        Return 0

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CreateFromLegacyXmlFile(lpFilePath As String) As FileDefinition
      Try
        If Not File.Exists(lpFilePath) Then
          Throw New ApplicationException("Invalid Path")
        End If

        Dim lobjDocument As New XmlDocument
        lobjDocument.Load(lpFilePath)

        Dim lobjFileDefinition As FileDefinition = CreateFromLegacyXml(lobjDocument)
        Dim lstrExtension As String = lobjFileDefinition.FileInfo.Extension
        If Not String.IsNullOrEmpty(lstrExtension) Then
          Dim lstrMimeType As String = Helper.GetMIMEType(lstrExtension.ToLower())
          ' First look in the registry
          If Not String.IsNullOrEmpty(lstrMimeType) Then
            lobjFileDefinition.FileInfo.MimeType = lstrMimeType
            'Else
            '  ' Then ask Microsoft
            '  Dim lstrMimeMap As String = Web.MimeMapping.GetMimeMapping(String.Format("a.{0}", lstrExtension))
            '  If lstrMimeMap <> "application/octet-stream" Then
            '    lobjFileDefinition.FileInfo.MimeType = lstrMimeMap
            '  End If
          End If
        End If

        Return lobjFileDefinition

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function CreateFromLegacyXml(lpXml As XmlDocument) As FileDefinition
      Try

        Dim lobjFileDefinition As New FileDefinition
        Dim lobjNode As XmlNode = lpXml.SelectSingleNode("//Info/FileType")
        Dim lobjPattern As HeaderPattern

        If lobjNode Is Nothing Then
          Throw New InvalidOperationException("Missing FileType node")
        End If
        lobjFileDefinition.FileInfo.FileType = lobjNode.InnerText

        lobjNode = lpXml.SelectSingleNode("//Info/Ext")
        'If lobjNode Is Nothing Then
        '  Throw New InvalidOperationException("Missing extension node")
        'End If
        If lobjNode IsNot Nothing Then
          lobjFileDefinition.FileInfo.Extension = lobjNode.InnerText.ToLower()
        End If

        lobjNode = lpXml.SelectSingleNode("//ExtraInfo/Rem")
        If lobjNode IsNot Nothing Then
          lobjFileDefinition.FileInfo.Remarks = lobjNode.InnerText
        End If

        lobjNode = lpXml.SelectSingleNode("//ExtraInfo/RefURL")
        If lobjNode IsNot Nothing Then
          lobjFileDefinition.FileInfo.RefUrl = lobjNode.InnerText
        End If

        Dim lobjNodeList As XmlNodeList = lpXml.SelectNodes("//FrontBlock/Pattern")
        For Each lobjNode In lobjNodeList
          lobjPattern = New HeaderPattern With {
            .Position = IntegerType.FromString(lobjNode.SelectSingleNode("Pos").InnerText),
            .Bytes = lobjNode.SelectSingleNode("Bytes").InnerText
          }
          lobjFileDefinition.FrontBlockSize = (lobjPattern.Position + lobjPattern.Length)
          lobjFileDefinition.Patterns.Add(lobjPattern)
        Next

        lobjNodeList = lpXml.SelectNodes("//GlobalStrings/String")
        For Each lobjNode In lobjNodeList
          Dim lstrByteString As ByteString = ByteStringFactory.Create(lobjNode.InnerText)
          lobjFileDefinition.GlobalStrings.Add(lstrByteString)
        Next

        Return lobjFileDefinition

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function FromJsonFile(lpFilePath As String) As FileDefinition
      Try
        Dim lstrJson As String = Helper.ReadAllTextFromFile(lpFilePath)
        Return FromJson(lstrJson)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function FromJson(lpJsonString As String) As FileDefinition
      Try
        Dim lobjJsonSettings As New JsonSerializerSettings()
        With lobjJsonSettings
          .TypeNameHandling = TypeNameHandling.None
        End With

        Return JsonConvert.DeserializeObject(lpJsonString, GetType(FileDefinition), lobjJsonSettings)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub SaveToJson(lpFilePath As String, Optional lpIndented As Boolean = False)
      Try
        Dim lobjJsonSerializer As New JsonSerializer

        With lobjJsonSerializer
          .TypeNameHandling = TypeNameHandling.None
          If lpIndented Then
            .Formatting = System.Xml.Formatting.Indented
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

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        If mobjFileInfo IsNot Nothing Then
          lobjIdentifierBuilder.Append(mobjFileInfo.DebuggerIdentifier)
        Else
          lobjIdentifierBuilder.Append("No File Info")
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

  End Class

End Namespace