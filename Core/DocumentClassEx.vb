'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core.ChoiceLists
Imports Documents.Search
Imports Documents.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Ionic.Zip

#End Region

Namespace Core

  <XmlInclude(GetType(TemplatedDocumentClass))>
  Partial Public Class DocumentClass
    Implements IComparable
    Implements ISerialize
    Implements IZipStreamSerialize
    Implements IXmlSerializable
    Implements ITableable

#Region "Constructors"

    ''' <summary>
    ''' Constructs a new document class object by deserializing from the XML file
    ''' referenced in the specified XML file path.
    ''' </summary>
    Public Sub New(ByVal lpXMLFilePath As String)
      Try
        Dim lobjDocumentClass As DocumentClass = Deserialize(lpXMLFilePath)
        Helper.AssignObjectProperties(lobjDocumentClass, Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DocumentClass::New('{0}')", lpXMLFilePath))
        Helper.DumpException(ex)

      End Try

    End Sub

    ''' <summary>
    ''' Constructs a new document class object using the specified stream and zip file object.
    ''' </summary>
    ''' <param name="lpStream">An IO.Stream object derived from a document class file.</param>
    ''' <param name="lpZipFile">A ZipFile object reference obtained during the deserialization of the parent repository file.</param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal lpStream As IO.Stream, ByVal lpZipFile As ZipFile)
      Try
        Dim lobjDocumentClass As DocumentClass = Deserialize(lpStream, lpZipFile)
        Helper.AssignObjectProperties(lobjDocumentClass, Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DocumentClass::New('{0}')", lpStream))
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Function GetAllChoiceLists() As ChoiceLists.ChoiceLists
      Try

        Return Me.Properties.GetAllChoiceLists

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      Try
        If TypeOf obj Is DocumentClass Then
          Return Name.CompareTo(obj.Name)
        Else
          Throw New ArgumentException("Object is not a DocumentClass")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DocumentClass::CompareTo('{0}')", obj.ToString))
      End Try

    End Function

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
        Return DOCUMENT_CLASS_FILE_EXTENSION
      End Get
    End Property

    Public Overloads Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Overloads Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.DeSerialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Return Serializer.Serialize.Xml(Me)
    End Function

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DocumentClass::Serialize('{0}', '{1}', '{2}')", lpFilePath, lpWriteProcessingInstruction, lpStyleSheetPath))
      End Try
    End Sub

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

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try
        If lpFileExtension.Length = 0 Then
          ' No override was provided
          If lpFilePath.EndsWith(DOCUMENT_CLASS_FILE_EXTENSION) = False Then
            lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & DOCUMENT_CLASS_FILE_EXTENSION
          End If
        End If
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Serialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpFileExtension))
      End Try
    End Sub

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return Me.GetType.FullName
      End Try
    End Function

#End Region

#Region "IZipStreamSerialize Implementation"

    Public Overloads Function Deserialize(ByVal lpStream As System.IO.Stream, ByVal lpZipFile As ZipFile) As Object Implements IZipStreamSerialize.Deserialize
      Try
        Return Serializer.Deserialize.FromZippedStream(lpStream, lpZipFile, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        Return Nothing
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

        If reader.GetType.Name <> "EnhancedStreamReader" AndAlso reader.GetType.IsSubclassOf(GetType(EnhancedStreamReader)) Then
          ApplicationLogging.WriteLogEntry("Unable to provide enhanced deserialization for object")
        End If

        Dim lstrFileName As String = String.Empty
        Dim lstrSerializationFolderPath As String = String.Empty

        If Helper.CallStackContainsMethodName("FromStream", "FromZippedStream") Then
          mobjSerializationSourceType = SourceType.Stream
        End If

        If mobjSerializationSourceType = SourceType.File Then
          If reader.BaseURI.Length > 0 Then
            ' Get the file name from the Base URI
            lstrFileName = reader.BaseURI.Replace("file:///", "").Replace("/", "\")
            ' Get the file folder
            lstrSerializationFolderPath = Path.GetDirectoryName(lstrFileName)
          Else
            ' The base URI is not set, the source must actually be a stream
            mobjSerializationSourceType = SourceType.Stream
          End If
        End If

        Dim lstrChoiceListFileName As String = String.Empty
        Dim lstrChoiceListFile As String = String.Empty
        Dim lobjClassificationProperty As ClassificationProperty = Nothing
        Dim lobjPropertyElement As XmlElement
        'Dim lobjTestElement As XmlElement
        Dim lobjChoiceListElement As XmlElement
        Dim lobjChoiceList As ChoiceList = Nothing

        Dim lobjXmlDocument As New XmlDocument

        Dim lobjTargetStream As IO.Stream = Nothing
        Dim lobjZipFile As ZipFile = Nothing
        Dim lobjChoiceListEntry As ZipEntry = Nothing

        lobjXmlDocument.Load(reader)

        With lobjXmlDocument

          ' Get the name
          Me.Name = .DocumentElement.GetAttribute("Name")

          ' Get the Id
          Me.ID = .DocumentElement.Item("ID").InnerText

          ' Get the Label
          Me.Label = .DocumentElement.Item("Label").InnerText

          ' Read the properties
          'Dim lobjPropertyNodes As XmlNodeList = .SelectNodes("//ClassificationProperty")
          Dim lobjPropertyNodes As XmlNodeList = .SelectNodes("//Properties/*")

          For Each lobjPropertyNode As XmlNode In lobjPropertyNodes

            lobjPropertyElement = CType(lobjPropertyNode, XmlElement)

            ' Create a new classification property
            'lobjClassificationProperty = New ClassificationProperty

            'lobjClassificationProperty = ClassificationPropertyFactory.Create( _
            '                               System.Enum.Parse(GetType(PropertyType), lobjPropertyElement.Item("Type").InnerText), _
            '                               lobjPropertyElement.Item("Name").InnerText, _
            '                               lobjPropertyElement.Item("SystemName").InnerText, _
            '                               System.Enum.Parse(GetType(Cardinality), lobjPropertyElement.Item("Cardinality").InnerText))

            'With lobjClassificationProperty
            '  '.Type = System.Enum.Parse(GetType(PropertyType), lobjPropertyElement.Item("Type").InnerText)
            '  '.Name = lobjPropertyElement.Item("Name").InnerText
            '  '.SystemName = lobjPropertyElement.Item("SystemName").InnerText
            '  '.Cardinality = System.Enum.Parse(GetType(Cardinality), lobjPropertyElement.Item("Cardinality").InnerText)
            '  .IsInherited = lobjPropertyElement.Item("IsInherited").InnerText
            '  .IsRequired = lobjPropertyElement.Item("IsRequired").InnerText
            '  .IsSystemProperty = lobjPropertyElement.Item("IsSystemProperty").InnerText
            '  .Searchable = lobjPropertyElement.Item("Searchable").InnerText

            '  ' Check to see if the Settability element exists
            '  lobjTestElement = lobjPropertyElement.Item("Settability")

            '  If lobjTestElement IsNot Nothing Then
            '    .Settability = System.Enum.Parse(GetType(ClassificationProperty.SettabilityEnum), lobjTestElement.InnerText)
            '  End If

            '  .Selectable = lobjPropertyElement.Item("Selectable").InnerText

            lobjClassificationProperty = ClassificationProperty.CreateFromXmlelement(lobjPropertyElement)

            ' Get the ChoiceList
            lobjChoiceListElement = lobjPropertyElement.Item("ChoiceList")

            If lobjChoiceListElement IsNot Nothing Then

              ' Get the name of the ChoiceList file
              lstrChoiceListFileName = lobjChoiceListElement.GetAttribute("href")

              If mobjSerializationSourceType = SourceType.File Then

                ' Build the fully qualified path
                lstrChoiceListFile = String.Format("{0}{1}", lstrSerializationFolderPath,
                                                      lstrChoiceListFileName)

                ' Construct a new ChoiceList file using the xml file path
                lobjChoiceList = New ChoiceList(lstrChoiceListFile)

              Else
                ' Get the choice list file out of the zip file
                If TypeOf (reader) Is XmlZipStreamReader Then
                  lobjZipFile = CType(reader, XmlZipStreamReader).ZipFile
                  lobjTargetStream = New MemoryStream


                  If lobjZipFile.ContainsEntry(lstrChoiceListFileName) = False Then
                    ' Make sure we have a clean file name
                    lstrChoiceListFileName = Helper.CleanFile(lstrChoiceListFileName, "-").Replace("\", String.Empty)
                  End If

                  ' Try one more time
                  If lobjZipFile.ContainsEntry(lstrChoiceListFileName) = False Then
                    Dim lstrChoiceListName As String = lobjChoiceListElement.GetAttribute("name")
                    Throw New Exceptions.DocumentClassNotInitializedException(
                      String.Format("Unable to extract choice list class file '{0}' from rfa file.  No entry exists by that name.",
                                    lstrChoiceListName), lstrChoiceListName, Nothing)
                  End If

                  lobjChoiceListEntry = lobjZipFile.Item(lstrChoiceListFileName)
                  lobjChoiceListEntry.Extract(lobjTargetStream)
                  lobjChoiceList = New ChoiceList(lobjTargetStream)

                End If
              End If

              ' Add the ChoiceList to the property
              If lobjChoiceList IsNot Nothing Then
                lobjClassificationProperty.ChoiceList = lobjChoiceList
              End If


            End If

            'End With

            ' Add the property to the collection
            Me.Properties.Add(lobjClassificationProperty)

          Next


        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        If writer.GetType.Name <> "EnhancedXmlTextWriter" AndAlso writer.GetType.IsSubclassOf(GetType(EnhancedXmlTextWriter)) Then
          ApplicationLogging.WriteLogEntry("Unable to provide enhanced serialization for object")
        End If

        Dim lstrFileName As String

        If (TypeOf (writer) Is EnhancedXmlTextWriter = False) AndAlso (Helper.CallStackContainsMethodName("ToStream")) Then
          If RepositoryName IsNot Nothing AndAlso RepositoryName.Length > 0 Then
            lstrFileName = String.Format("{0}_{1}.{2}", RepositoryName, Me.Name, DocumentClass.DOCUMENT_CLASS_FILE_EXTENSION)
            lstrFileName = Helper.CleanFile(lstrFileName, "-")
          Else
            Throw New Exceptions.RepositoryNotAvailableException(String.Empty,
              String.Format("Unable to write xml for document class '{0}', the repository name is note available for the document class file name.", Me.Name))
          End If
        Else
          lstrFileName = CType(writer, EnhancedXmlTextWriter).FileName
        End If

        Dim lstrFileLocation As String = IO.Path.GetDirectoryName(lstrFileName)
        Dim lstrChoiceListFileName As String = String.Empty
        Dim lobjChoiceList As ChoiceLists.ChoiceList = Nothing

        With writer

          ' Write a human readable file
          '.Formatting = Formatting.Indented

          ' Write the Document Class Name element
          .WriteAttributeString("Name", Me.Name)

          '.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
          .WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")
          '.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
          .WriteAttributeString("xmlns", "xsd", Nothing, "http://www.w3.org/2001/XMLSchema")

          ' Write the ID element
          .WriteElementString("ID", ID)

          ' Write the Label element
          .WriteElementString("Label", Label)

          ' Open the Properties Element
          .WriteStartElement("Properties")

          For Each lobjProperty As ClassificationProperty In Me.Properties

            If lobjProperty.ChoiceList Is Nothing Then
              ' Write the ClassificationProperty element
              .WriteRaw(lobjProperty.ToXmlString)
            Else
              ' We want to write the Choice list to a separate file

              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              '
              ' NOTE: We are writing each property out here manually.
              '       It is imperative that with any enhancements to the 
              '       ClassificationProperty object model this section is updated.
              '       If not, then the new properties will not be serialized 
              '       out if the property contains a choice list.
              ' 
              ' Ernie Bahr
              ' 3/22/2010
              '
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

              ' Open the ClassificationProperty Element
              .WriteStartElement(lobjProperty.GetType.Name)

              ' Write the attributes
              .WriteAttributeString("ID", lobjProperty.ID)
              .WriteAttributeString("Description", lobjProperty.Description)

              ' Write all the sub elements
              .WriteElementString("Type", lobjProperty.Type.ToString)
              .WriteElementString("Cardinality", lobjProperty.Cardinality.ToString)
              .WriteElementString("Name", lobjProperty.Name)
              .WriteElementString("SystemName", lobjProperty.SystemName)
              If lobjProperty.DefaultValue IsNot Nothing Then
                .WriteElementString("DefaultValue", lobjProperty.DefaultValue.ToString)
              Else
                .WriteElementString("DefaultValue", String.Empty)
              End If
              .WriteElementString("IsInherited", lobjProperty.IsInherited.ToString.ToLower)
              .WriteElementString("IsRequired", lobjProperty.IsRequired.ToString.ToLower)
              .WriteElementString("IsHidden", lobjProperty.IsHidden.ToString.ToLower)
              .WriteElementString("IsSystemProperty", lobjProperty.IsSystemProperty.ToString.ToLower)
              .WriteElementString("Searchable", lobjProperty.Searchable.ToString.ToLower)
              .WriteElementString("Settability", lobjProperty.Settability.ToString)
              .WriteElementString("Selectable", lobjProperty.Selectable.ToString.ToLower)
              .WriteElementString("Value", lobjProperty.Value)
              ' TODO: Write out the values .WriteElementString("Values", lobjProperty.Type.ToString)

              ' Get the information specific to the particular classification property type
              Dim lobjPropertyAttributes() As PropertyInfo = lobjProperty.GetType.GetProperties
              Dim lobjPropertyAttributeValue As Object

              For Each lobjPropertyInfo As PropertyInfo In lobjPropertyAttributes
                If lobjPropertyInfo.CanRead AndAlso lobjPropertyInfo.CanWrite Then
                  Select Case lobjPropertyInfo.Name
                    Case "MaxLength", "MinValue", "MaxValue", "DefaultValueString"
                      lobjPropertyAttributeValue = lobjPropertyInfo.GetValue(lobjProperty, Nothing)
                      .WriteElementString(lobjPropertyInfo.Name, lobjPropertyAttributeValue)
                  End Select
                End If
              Next

              ' Write the ChoiceList information
              lobjChoiceList = lobjProperty.ChoiceList

              ' Open the DocumentClasses Element
              .WriteStartElement("ChoiceList")

              ' Write the name of the ChoiceList as an attribute
              .WriteAttributeString("name", lobjChoiceList.Name)

              ' Write the location of the ChoiceList file
              If String.IsNullOrEmpty(RepositoryName) Then
                lstrChoiceListFileName = String.Format("{0}.{1}",
                                                          lobjChoiceList.Name,
                                                          ChoiceLists.ChoiceList.CHOICE_LIST_FILE_EXTENSION)
              Else
                lstrChoiceListFileName = String.Format("{0}_{1}.{2}",
                                                          RepositoryName, lobjChoiceList.Name,
                                                          ChoiceLists.ChoiceList.CHOICE_LIST_FILE_EXTENSION)
              End If

              ' It is possible that there were characters used in the 
              ' choice list name that are not valid for a file name.
              ' Scrub any invalid file name characters from the file name
              lstrChoiceListFileName = Helper.CleanFile(lstrChoiceListFileName, "-")

              .WriteAttributeString("href", lstrChoiceListFileName)

              If Repository IsNot Nothing AndAlso Repository.ChildStreams IsNot Nothing Then
                ' Persist the choice list in the collection of child streams
                Repository.ChildStreams.Add(New ZipFileStream(lstrChoiceListFileName, String.Empty, lobjChoiceList.SerializeToStream))
              Else
                ' Serialize the choice list file
                lstrChoiceListFileName = String.Format("{0}\{1}", lstrFileLocation, lstrChoiceListFileName)
                lobjChoiceList.Serialize(lstrChoiceListFileName)
              End If

              ' Close the ChoiceList element
              .WriteEndElement()

              .WriteElementString("SubscribedClasses", String.Empty)

              ' Close the ClassificationProperty element
              .WriteEndElement()

            End If

          Next

          ' Close the Properties element
          .WriteEndElement()

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "ITableable Implementation"

    Public Function ToDataTable() As System.Data.DataTable Implements ITableable.ToDataTable
      Try
        Dim lobjDataTable As DataTable = Me.Properties.ToDataTable
        lobjDataTable.TableName = String.Format("tbl{0}", Me.GetType.Name)
        Return lobjDataTable
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
