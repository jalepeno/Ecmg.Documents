' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Parameter.vb
'  Description :  [type_description_here]
'  Created     :  11/17/2011 4:17:03 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core

  <Category("Behavior"),
  JsonConverter(GetType(ParameterJsonConverter))>
  Public Class Parameter
    Inherits ECMProperty
    Implements IParameter
    Implements IJsonSerializable(Of IParameter)
    Implements IXmlSerializable

#Region "Class Constants"

    Public Const PARAMETER_EXPRESSION As String = "(?<Prefix>.*){(?<ParamName>[a-zA-Z0-9]*):(?<ParamValue>[a-zA-Z0-9]*)}(?<Suffix>.*)"

#End Region

#Region "Constructors"

    Protected Friend Sub New()
      MyBase.New(String.Empty)
    End Sub

    ''' <summary>
    ''' Intended for internal use only by the framework
    ''' </summary>
    ''' <param name="lpInternalUseOnly"></param>
    ''' <remarks></remarks>
    Protected Friend Sub New(ByVal lpInternalUseOnly As String)
      MyBase.New(lpInternalUseOnly)
    End Sub

    Protected Friend Sub New(ByVal lpMenuType As PropertyType,
                 ByVal lpName As String,
                 ByVal lpValue As Object)
      MyBase.New(String.Empty)
      Try
        Type = lpMenuType
        Name = lpName
        Value = lpValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpParameter As IParameter)
      MyBase.New(String.Empty)
      Try
        Me.Type = lpParameter.Type
        Me.Name = lpParameter.Name
        Me.SystemName = lpParameter.SystemName
        Me.Cardinality = lpParameter.Cardinality
        Me.Description = lpParameter.Description
        Me.SetPersistence(lpParameter.Persistent)
        Select Case lpParameter.Cardinality
          Case Core.Cardinality.ecmSingleValued
            Me.Value = lpParameter.Value
            Me.DefaultValue = lpParameter.DefaultValue
          Case Core.Cardinality.ecmMultiValued
            Me.Values = lpParameter.Values
        End Select
        If Me.Type = PropertyType.ecmEnum AndAlso Me.Cardinality = Cardinality.ecmSingleValued Then
          Me.EnumType = CType(lpParameter, SingletonEnumParameter).EnumType
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates a parameter from an xml node
    ''' </summary>
    ''' <param name="lpXmlNode">An xml node from a serialized object containing the parameter.</param>
    ''' <remarks></remarks>
    Friend Sub New(lpXmlNode As XmlNode)
      MyBase.New(String.Empty)
      Try
        InitializeFromXmlNode(lpXmlNode)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Shared Methods"

    ''' <summary>
    ''' Creates a parameter from an xml node
    ''' </summary>
    ''' <param name="lpXmlNode">An xml node from a serialized object containing the parameter.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Create(lpXmlNode As XmlNode) As IParameter
      Try

        Dim lstrParameterType As String = String.Empty
        Dim lobjParameterType As Type = Nothing
        Dim lobjParameter As IParameter = Nothing

        lstrParameterType = lpXmlNode.Name
        If Not String.IsNullOrEmpty(lstrParameterType) Then
          If String.Compare("Parameter", lstrParameterType) = 0 Then
            lobjParameter = New Parameter(lpXmlNode)
          Else
            lobjParameterType = Helper.GetTypeFromAssembly(Reflection.Assembly.GetExecutingAssembly, lstrParameterType)
            lobjParameter = Activator.CreateInstance(lobjParameterType)
            lobjParameter.InitializeFromXmlNode(lpXmlNode)
          End If
        End If

        Return ParameterFactory.Create(lobjParameter)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetInlineParameter(ByVal lpValue As String) As INameValuePair
      Try

        Dim lobjRegex As Regex = New Regex(PARAMETER_EXPRESSION,
            RegexOptions.CultureInvariant Or RegexOptions.Compiled)

        ' Split the InputText wherever the regex matches
        Dim lstrResults As String() = lobjRegex.Split(lpValue)

        ' Test to see if there is a match in the InputText
        Dim lblnIsMatch As Boolean = lobjRegex.IsMatch(lpValue)

        If lblnIsMatch Then
          Dim lintParameterNameGroupNumber As Integer = lobjRegex.GroupNumberFromName("ParamName")
          Dim lintParameterValueGroupNumber As Integer = lobjRegex.GroupNumberFromName("ParamValue")

          If lintParameterNameGroupNumber > 0 AndAlso lintParameterValueGroupNumber > 0 Then
            Return New Configuration.KeyValuePair(lstrResults(lintParameterNameGroupNumber),
                                                  lstrResults(lintParameterValueGroupNumber))
          Else
            Return Nothing
          End If

        Else
          Return Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


    Public Shadows Function ToXmlString() As String Implements IParameter.ToXmlString
      Try
        If Me.Type <> PropertyType.ecmObject Then
          Return MyBase.ToXmlString
        Else
          Return Serializer.Serialize.XmlElementString(Me)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ICloneable Implementation"

    Public Overloads Function Clone() As Object
      Try
        Return New Parameter(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IJsonSerializable(Of Parameter)"

    Public Overrides Function ToJson() As String Implements IJsonSerializable(Of IParameter).ToJson
      Try
        Return JsonConvert.SerializeObject(Me, Newtonsoft.Json.Formatting.None, New ParameterJsonConverter())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function FromJson(lpJson As String) As IParameter Implements IJsonSerializable(Of IParameter).FromJson
      Try
        Return JsonConvert.DeserializeObject(lpJson, GetType(IParameter), New ParameterJsonConverter())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function CreateFromJsonReader(reader As JsonReader) As Parameter
      Try
        'Return JsonConvert.DeserializeObject(reader, GetType(IOperation), New OperationJsonConverter())
        Dim lobjConverter As New ParameterJsonConverter()
        Return lobjConverter.ReadJson(reader, Nothing, Nothing, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Overrides"

    Public Overrides Sub ReadXml(ByVal reader As System.Xml.XmlReader)

      ' NOTE: This method is not complete

      Dim lobjXmlDocument As New XmlDocument

      Try

        lobjXmlDocument.Load(reader)

        With lobjXmlDocument

          ' Get the name
          Me.Name = .DocumentElement.GetAttribute("Name")

          Me.Value = .DocumentElement.GetAttribute("Value")

          Dim lstrCardinality As String = .DocumentElement.GetAttribute("Cardinality")

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

          ' Write the DisplayName attribute
          .WriteAttributeString("DisplayName", DisplayName)

          ' Write the Type attribute
          .WriteAttributeString("Type", Type.ToString)

          If Me.Type = PropertyType.ecmEnum Then
            .WriteAttributeString("EnumType", Me.EnumTypeName)
          End If

          ' Write the Description attribute
          .WriteAttributeString("Description", Description)

          If Cardinality = Core.Cardinality.ecmSingleValued Then
            ' Write the Type attribute
            If Value IsNot Nothing Then
              .WriteAttributeString("Value", Value.ToString)
            Else
              .WriteAttributeString("Value", String.Empty)
            End If
          Else
            ' Write the Cardinality attribute
            .WriteAttributeString("Cardinality", Cardinality.ToString)

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

    Friend Overridable Sub InitializeFromXmlNode(lpXmlNode As XmlNode) Implements IParameter.InitializeFromXmlNode
      Try

        If lpXmlNode Is Nothing Then
          Throw New ArgumentNullException("lpXmlNode")
        End If

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

        ' Get the display name
        lobjAttribute = lpXmlNode.Attributes("DisplayName")
        If lobjAttribute IsNot Nothing Then
          DisplayName = lobjAttribute.InnerText
        End If

        ' Get the Description
        lobjAttribute = lpXmlNode.Attributes("Description")
        If lobjAttribute IsNot Nothing Then
          Description = lobjAttribute.InnerText
        End If

        ' Get the Type
        lobjAttribute = lpXmlNode.Attributes("Type")
        If lobjAttribute Is Nothing Then
          Throw New ArgumentException("Node has no Type attribute")
        End If
        Type = [Enum].Parse(GetType(PropertyType), lobjAttribute.InnerText)

        If Type = PropertyType.ecmEnum Then
          lobjAttribute = lpXmlNode.Attributes("EnumType")
          If lobjAttribute Is Nothing Then
            Throw New ArgumentException("Node has no EnumType attribute")
          End If
          ' <Modified by: Ernie at 4/7/2014-2:30:49 PM on machine: ERNIE-THINK>
          ' This modification is for backwards compatibility.  
          ' The enum was originally called SaveMode but it was renamed to SaveModeEnum.
          If lobjAttribute.InnerText = "SaveMode" Then
            Me.EnumTypeName = "SaveModeEnum"
          Else
            Me.EnumTypeName = lobjAttribute.InnerText
          End If
          ' </Modified by: Ernie at 4/7/2014-2:30:49 PM on machine: ERNIE-THINK>

          ' Me.EnumType = Helper.GetTypeFromAssembly(Reflection.Assembly.GetExecutingAssembly, EnumTypeName)
          'If Me.EnumType IsNot Nothing Then
          '  Dim lobjEnumDictionary As IDictionary(Of String, Integer) = Helper.EnumerationDictionary(Me.EnumType)
          '  mobjStandardValues = lobjEnumDictionary.Keys
          'End If

        End If

        ' Get the Value
        lobjAttribute = lpXmlNode.Attributes("Value")
        If lobjAttribute IsNot Nothing Then
          Value = lobjAttribute.InnerText
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