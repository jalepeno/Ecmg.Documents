' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ObjectParameter.vb
'  Description :  [type_description_here]
'  Created     :  8/14/2012 1:36:46 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ObjectParameter
    Inherits Parameter

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpValue As Object,
               ByVal lpDescription As String)
      MyBase.New(lpMenuType, lpName, lpValue)
      Description = lpDescription
    End Sub

    Friend Sub New(lpXmlNode As XmlNode)
      MyBase.New()
      Try
        InitializeFromXmlNode(lpXmlNode)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub New(lpParameter As ObjectParameter)
      MyBase.New(lpParameter)
      Try
        Value = lpParameter.Value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ICloneable Implementation"

    Public Overloads Function Clone() As Object
      Try
        Return New ObjectParameter(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Overrides"

    Public Overrides Sub ReadXml(ByVal reader As System.Xml.XmlReader)
      Dim lobjXmlDocument As New XmlDocument

      Try

        lobjXmlDocument.Load(reader)

        With lobjXmlDocument

          ' Get the name
          Me.Name = .DocumentElement.GetAttribute("Name")
          Dim lstrCardinality As String = .DocumentElement.GetAttribute("Cardinality")

          ' NOTE: THE IMPLEMENTATION BELOW ASSUMES THE VALUE IS SINGLE_VALUED
          ' Ernie Bahr - 8/14/2012
          Dim lstrObjectType As String = .DocumentElement.Item("Value").FirstChild.Name
          Dim lobjObjectType As Type = System.Type.GetType(lstrObjectType)
          Dim lstrValueXml As String = .DocumentElement.Item("Value").FirstChild.InnerXml
          Dim lobjValue As Object = Serializer.Deserialize.XmlString(lstrValueXml, lobjObjectType)

          Me.Value = .DocumentElement.GetAttribute("Value")

        End With

        Type = [Enum].Parse(GetType(PropertyType), reader.GetAttribute("Type"))

        reader.GetAttribute("Value")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Sub WriteXml(ByVal writer As System.Xml.XmlWriter)
      Try

        With writer

          ' Write the Name attribute
          .WriteAttributeString("Name", Name)

          ' Write the Type attribute
          .WriteAttributeString("Type", Type.ToString)

          If Cardinality = Core.Cardinality.ecmSingleValued Then
            .WriteStartElement("Value")

            If Value IsNot Nothing Then
              .WriteRaw(Serializer.Serialize.XmlElementString(Value))
            End If

            .WriteEndElement()
          Else
            ' Write the Cardinality attribute
            .WriteAttributeString("Cardinality", Cardinality.ToString)

            ' NOTE: STILL NEED TO UPDATE VALUES TO THE COMPLEX FORM
            If Values.Count = 0 Then
              .WriteElementString("Values", Nothing)
            Else
              .WriteStartElement("Values")
              For Each lobjValue As Object In Values
                .WriteElementString("Value", lobjValue.ToString)
              Next
              .WriteEndElement()
            End If
          End If

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Friend Methods"

    Friend Overrides Sub InitializeFromXmlNode(lpXmlNode As XmlNode)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpXmlNode)
#Else
          If lpXmlNode Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpXmlNode))
          End If
#End If
        'If lpXmlNode.HasChildNodes = False Then
        '  Throw New ArgumentException("Node has no child nodes", "lpXmlNode")
        'End If

        Dim lobjNode As XmlNode = Nothing

        Dim lobjAttribute As XmlAttribute = Nothing

        ' Get the Name
        lobjAttribute = lpXmlNode.Attributes("Name")
        If lobjAttribute Is Nothing Then
          Throw New ArgumentException("Node has no name attribute")
        End If
        Name = lobjAttribute.InnerText


        ' Get the Type
        lobjAttribute = lpXmlNode.Attributes("Type")
        If lobjAttribute Is Nothing Then
          Throw New ArgumentException("Node has no Type attribute")
        End If
        Type = [Enum].Parse(GetType(PropertyType), lobjAttribute.InnerText)

        ' Get the Value
        lobjNode = lpXmlNode.Item("Value")
        If lobjNode IsNot Nothing Then
          Dim lstrObjectType As String = lobjNode.FirstChild.Name
          Dim lobjObjectType As Type = Helper.GetTypeFromAssembly(Reflection.Assembly.GetExecutingAssembly, lstrObjectType)
          If lobjObjectType Is Nothing Then
            lobjObjectType = Helper.GetTypeFromAssembly(Reflection.Assembly.GetEntryAssembly, lstrObjectType, True)
            If lobjObjectType Is Nothing Then
              Throw New Exceptions.ItemDoesNotExistException(lstrObjectType)
            End If

          End If
          Dim lstrValueXml As String = lobjNode.InnerXml
          Dim lobjValue As Object = Serializer.Deserialize.XmlString(lstrValueXml, lobjObjectType)
          Value = lobjValue
        Else
          ' If we have a value attribute then it is a 
          ' single valued parameter and we do not need 
          ' to continue, since we don't we need to check 
          ' for the Cardinality and Values info.

          ' Get the Cardinality
          lobjAttribute = lpXmlNode.Attributes("Cardinality")
          If lobjAttribute IsNot Nothing Then
            Cardinality = [Enum].Parse(GetType(Cardinality), lobjAttribute.InnerText)
          End If

          If Cardinality = Core.Cardinality.ecmMultiValued Then
            lobjNode = lpXmlNode.SelectSingleNode("Values")
            If lobjNode IsNot Nothing AndAlso lobjNode.HasChildNodes Then
              For Each lobjChildNode As XmlNode In lobjNode.ChildNodes
                Values.Add(lobjChildNode.InnerText)
              Next
            End If
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace