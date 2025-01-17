'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities


#End Region

Namespace Comparison

  Partial Public Class DocumentComparison
    Implements ISerialize
    Implements IXmlSerializable

#Region "Class Variables"

    Private mstrXsiType As String = String.Empty
    Private mstrDocumentXID As String = String.Empty
    Private mstrDocumentYID As String = String.Empty

#End Region

#Region "ISerialize Implementation"

    Public ReadOnly Property DefaultFileExtension As String Implements ISerialize.DefaultFileExtension
      Get
        Return DOCUMENT_COMPARISON_FILE_EXTENSION
      End Get
    End Property

    Public ReadOnly Property DefaultArchiveFileExtension As String
      Get
        Return DOCUMENT_COMPARISON_ARCHIVE_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Dim lstrExtension As String = IO.Path.GetExtension(lpFilePath)

        Select Case lstrExtension.TrimStart(".")
          Case DefaultArchiveFileExtension
            Return FromArchive(lpFilePath)
          Case DefaultFileExtension
            Dim lobjDocumentComparison As DocumentComparison = Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)

            Return lobjDocumentComparison
          Case Else
            Throw New InvalidOperationException(
              String.Format("Extension '{0}' not expected.  Deserialization only allowed from extensions '{1}' and '{2}'",
                            lstrExtension, DefaultArchiveFileExtension, DefaultFileExtension))
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function Deserialize(lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Serializer.Serialize.Xml(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Serialize(ByRef lpFilePath As String, lpFileExtension As String) Implements ISerialize.Serialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        If lpFileExtension.Length = 0 Then
          ' No override was provided
          If lpFilePath.EndsWith(DefaultFileExtension) = False Then
            lpFilePath = lpFilePath.Remove(lpFilePath.Length - DefaultFileExtension.Length) & DefaultFileExtension
          End If

        End If

        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(lpFilePath As String) Implements ISerialize.Serialize
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(lpFilePath As String, lpWriteProcessingInstruction As Boolean, lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Serializer.Serialize.XmlString(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      ' As per the Microsoft guidelines this is not implemented
      Return Nothing
    End Function

    Public Sub ReadXml(reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Try
        Try

          Dim lobjComparisonProperty As ComparisonProperty = Nothing
          Dim lobjPropertyX As IProperty = Nothing
          Dim lobjPropertyY As IProperty = Nothing
          Dim lenuPropertyType As PropertyType
          Dim lenuCardinality As Cardinality
          Dim lstrName As String = Nothing
          Dim lstrValue As String = Nothing
          Dim lobjValues As Object = Nothing
          Dim lstrScope As String = String.Empty
          Dim lenuScope As PropertyScope
          Dim lstrVersionId As String = String.Empty
          Dim lstrElementName As String = String.Empty
          Dim lobjPropertyValues As Values = Nothing

          Dim lblnInAllPropertiesNode As Boolean

          mstrXsiType = reader.GetAttribute("xsi:type")
          Dim lstrCurrentElementName As String = String.Empty

          Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name = "PropertiesInBothDocuments"
            If reader.Name = "AllProperties" Then
              If lblnInAllPropertiesNode = True Then
                ' We are leaving the 'AllProperties' node
                ' This is all we need
                reader.Close()
                Exit Sub
              End If
              lblnInAllPropertiesNode = Not lblnInAllPropertiesNode
            End If

            If lblnInAllPropertiesNode = False AndAlso reader.NodeType = XmlNodeType.Element Then
              Select Case reader.Name
                Case "DocumentX"
                  mstrDocumentXID = reader.GetAttribute("ID")
                Case "DocumentY"
                  mstrDocumentYID = reader.GetAttribute("ID")
              End Select
            End If

            If lblnInAllPropertiesNode = True Then
              Select Case reader.NodeType
                Case XmlNodeType.Element
                  lstrElementName = reader.Name
                  Select Case lstrElementName
                    Case "ComparisonProperty"
                      lstrScope = reader.GetAttribute("Scope")
                      If Not String.IsNullOrEmpty(lstrScope) Then
                        lenuScope = [Enum].Parse(lenuScope.GetType, lstrScope)
                      End If
                      lstrVersionId = reader.GetAttribute("VersionId")
                    Case "PropertyX"
                      lobjPropertyX = Nothing
                      lobjPropertyValues = New Values
                    Case "PropertyY"
                      lobjPropertyY = Nothing
                      lobjPropertyValues = New Values
                    Case "Values"
                      lobjPropertyValues = GetMultiValuedPropertyValues(reader.Name, reader)
                  End Select

                Case XmlNodeType.Text
                  Select Case lstrElementName
                    Case "Type"
                      lenuPropertyType = [Enum].Parse(lenuPropertyType.GetType, reader.Value)
                    Case "Cardinality"
                      lenuCardinality = [Enum].Parse(lenuCardinality.GetType, reader.Value)
                    Case "Name"
                      lstrName = reader.Value
                    Case "Value"
                      lstrValue = reader.Value
                  End Select

                Case XmlNodeType.EndElement
                  Select Case reader.Name
                    Case "ComparisonProperty"
                      lobjComparisonProperty = New ComparisonProperty(lobjPropertyX, lobjPropertyY)
                      lobjComparisonProperty.Scope = lenuScope
                      lobjComparisonProperty.VersionId = lstrVersionId
                      AllProperties.Add(lobjComparisonProperty)
                    Case "PropertyX"
                      lobjPropertyX = PropertyFactory.Create(lenuPropertyType, lstrName, lenuCardinality)
                      If lenuCardinality = Cardinality.ecmSingleValued Then
                        lobjPropertyX.Value = lstrValue
                      Else
                        lobjPropertyX.Values = lobjPropertyValues
                      End If
                    Case "PropertyY"
                      lobjPropertyY = PropertyFactory.Create(lenuPropertyType, lstrName, lenuCardinality)
                      If lenuCardinality = Cardinality.ecmSingleValued Then
                        lobjPropertyY.Value = lstrValue
                      Else
                        lobjPropertyY.Values = lobjPropertyValues
                      End If
                  End Select
              End Select
            End If
            reader.Read()
          Loop

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function GetMultiValuedPropertyValues(nodeName As String, reader As System.Xml.XmlReader) As Values
      Try
        Dim lobjReturnValues As New Values


        Do Until reader.Name = nodeName AndAlso reader.NodeType = XmlNodeType.EndElement
          reader.Read()
          Select Case reader.NodeType
            Case XmlNodeType.Element, XmlNodeType.EndElement
              Select Case reader.Name
                Case "Value", "anyType"
                  ' This is what we are looking for...
                Case "Values"
                  Exit Do
                Case Else
                  Exit Do
              End Select
            Case XmlNodeType.Text
              lobjReturnValues.Add(reader.Value)
          End Select
        Loop

        Return lobjReturnValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub WriteXml(writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        With writer

          ' Write the Name element
          .WriteAttributeString("Name", Me.Name)
          ' Write the Description element
          .WriteAttributeString("Description", Me.Description)

          .WriteAttributeString("TotalPropertyCount", Me.AllProperties.Count)

          .WriteAttributeString("PropertiesInBothDocuments", Me.PropertiesInBothDocuments.Count)
          .WriteAttributeString("PropertiesInXOnly", Me.PropertiesInXOnly.Count)
          .WriteAttributeString("PropertiesInYOnly", Me.PropertiesInYOnly.Count)
          .WriteAttributeString("Equal", Me.EqualProperties.Count)
          .WriteAttributeString("Unequal", Me.UnequalProperties.Count)

          .WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
          .WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")

          ' Write DocumentX
          WriteDocumentNode(writer, "DocumentX", DocumentX)

          ' Write DocumentY
          WriteDocumentNode(writer, "DocumentY", DocumentY)

          ' PLace all the property collections inside a parent Properties element
          .WriteStartElement("Properties")

          ' Write the AllProperties Element
          WritePropertyNode(writer, "AllProperties", Me.AllProperties)

          ' Write the AllProperties Element
          WritePropertyNode(writer, "PropertiesInBothDocuments", Me.PropertiesInBothDocuments)

          ' Write the AllProperties Element
          WritePropertyNode(writer, "PropertiesInXOnly", Me.PropertiesInXOnly)

          ' Write the AllProperties Element
          WritePropertyNode(writer, "PropertiesInYOnly", Me.PropertiesInYOnly)

          ' Write the AllProperties Element
          WritePropertyNode(writer, "EqualProperties", Me.EqualProperties)

          ' Write the AllProperties Element
          WritePropertyNode(writer, "UnequalProperties", Me.UnequalProperties)

          .WriteEndElement()

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub WriteDocumentNode(writer As System.Xml.XmlWriter, nodeName As String, document As Document)
      Try

        Dim lstrCPFName As String = Nothing

        With writer

          ' Open the Element
          .WriteStartElement(nodeName)
          If document IsNot Nothing Then
            ' Write the id as an attribute
            .WriteAttributeString("ID", document.ID)
            ' Write the name of the ChoiceList as an attribute
            .WriteAttributeString("Summary", document.DebuggerIdentifier)

            lstrCPFName = String.Format("{0}.{1}", document.ID, document.DefaultArchiveFileExtension)
            lstrCPFName = Helper.CleanFile(lstrCPFName, "-")

            .WriteAttributeString("href", lstrCPFName)
          End If

          ' Close the element
          .WriteEndElement()

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub WritePropertyNode(writer As System.Xml.XmlWriter, nodeName As String, properties As ComparisonProperties)
      Try
        With writer
          ' Open the element
          .WriteStartElement(nodeName)

          .WriteAttributeString("Count", properties.Count)

          ' Break them out into separate files
          For Each lobjProperty As IProperty In properties
            .WriteRaw(Serializer.Serialize.XmlElementString(lobjProperty))
          Next

          ' Close the element
          .WriteEndElement()

        End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace