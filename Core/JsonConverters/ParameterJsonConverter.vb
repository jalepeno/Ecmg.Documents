' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ParameterJsonConverter.vb
'  Description :  [type_description_here]
'  Created     :  01/22/2024 2:12:13 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.TypeConverters
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core

  Public Class ParameterJsonConverter
    Inherits JsonConverter

    Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
      Try
        Dim lobjParameter As Parameter = DirectCast(value, Parameter)
        Dim lenuEnumParameter As SingletonEnumParameter = Nothing

        With writer
          If Helper.IsRunningInstalled Then
            .Formatting = Formatting.None
          Else
            .Formatting = Formatting.Indented
          End If

          .WriteStartObject()

          ' Write the Parameter Type
          .WritePropertyName("type")
          .WriteValue(lobjParameter.GetType.Name)

          ' Write the 'Name' property
          .WritePropertyName("name")
          .WriteValue(lobjParameter.Name)

          ' Write the 'DisplayName' property
          .WritePropertyName("displayname")
          .WriteValue(lobjParameter.DisplayName)

          ' Write the 'Type' property
          .WritePropertyName("datatype")
          .WriteValue(EnumDescriptionTypeConverter.GetDescription(value.Type.GetType, value.Type.ToString))

          ' If the parameter type is an enum parameter, write the 'Enum Type' property
          If lobjParameter.Type = PropertyType.ecmEnum Then
            ' Write the 'Enum Type' property
            .WritePropertyName("enumtype")
            lenuEnumParameter = lobjParameter
            If Not String.IsNullOrEmpty(lenuEnumParameter.EnumTypeName) Then
              .WriteValue(lenuEnumParameter.EnumTypeName)
            Else
              .WriteValue(String.Empty)
            End If
          End If

          ' Write the 'Description' property
          .WritePropertyName("description")
          .WriteValue(lobjParameter.Description)

          ' Write the 'Value' property
          .WritePropertyName("value")

          Select Case lobjParameter.Type
            Case PropertyType.ecmBoolean
              .WriteValue(Boolean.Parse(lobjParameter.Value))
            Case PropertyType.ecmEnum
              .WriteValue(lobjParameter.ValueString)
            Case Else
              .WriteValue(lobjParameter.Value)
          End Select

          .WriteEndObject()

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function ReadJson(reader As JsonReader, objectType As Type, existingValue As Object, serializer As JsonSerializer) As Object
      Try

        Dim lstrCurrentPropertyName As String = String.Empty
        Dim lstrType As String = String.Empty
        Dim lstrName As String = String.Empty
        Dim lstrDisplayName As String = String.Empty
        Dim lstrDataType As String = String.Empty
        Dim lstrDescription As String = String.Empty
        Dim lobjValue As Object = Nothing
        Dim lstrEnumType As String = String.Empty
        Dim lenuPropertyType As PropertyType = Nothing
        Dim lobjEnumType As Type = Nothing
        Dim lobjParameters As New Parameters

        While reader.Read
          Select Case reader.TokenType
            Case JsonToken.PropertyName
              lstrCurrentPropertyName = reader.Value

            Case JsonToken.String, JsonToken.Boolean, JsonToken.Date, JsonToken.Integer, JsonToken.Float
              Select Case lstrCurrentPropertyName
                Case "type"
                  lstrType = reader.Value
                Case "name"
                  lstrName = reader.Value
                Case "displayname"
                  lstrDisplayName = reader.Value
                Case "datatype"
                  lstrDataType = reader.Value
                  'If lstrDataType = "Enumeration" Then
                  '  lenuDataType = [Enum].Parse(GetType(PropertyType), lstrDataType)
                  'End If
                  lenuPropertyType = EnumDescriptionTypeConverter.GetValue(GetType(PropertyType), lstrDataType)
                Case "enumtype"
                  lstrEnumType = reader.Value
                  lobjEnumType = Helper.GetTypeFromName(lstrEnumType)
                Case "description"
                  lstrDescription = reader.Value
                Case "value"
                  lobjValue = reader.Value
                Case "parameters"
                  lstrCurrentPropertyName = reader.Value
                  ' TODO: Handle the parameters

              End Select

            Case JsonToken.StartArray
              Select Case lstrCurrentPropertyName
                Case "parameters"
                  ' TODO: Handle the parameters
                  lobjParameters.Add(Parameter.CreateFromJsonReader(reader))
              End Select

            Case JsonToken.StartObject
              Select Case lstrCurrentPropertyName
                Case "parameters"
                  ' TODO: Handle the parameters

              End Select

            Case JsonToken.EndObject
              If lenuPropertyType = PropertyType.ecmEnum Then
                Return ParameterFactory.Create(lenuPropertyType, lstrName, lobjValue, lobjEnumType, lstrDescription)
              Else
                Return ParameterFactory.Create(lenuPropertyType, lstrName, lobjValue, lstrDescription)
              End If


          End Select
        End While

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function CanConvert(objectType As Type) As Boolean
      If TypeOf objectType Is IParameter OrElse Helper.IsAssignableFrom("IParameter", objectType) Then
        Return True
      Else
        Return False
      End If
      'Return objectType = GetType(Parameter)
    End Function

  End Class

End Namespace