' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  TransformationJsonConverter.vb
'  Description :  [type_description_here]
'  Created     :  01/21/2025 1:24:13 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------


#Region "Imports"

Imports System.Globalization
Imports Documents.Core
Imports Documents.Transformations
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region
Public Class TransformationJsonConverter
  Inherits JsonConverter

  Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
    Try
      Dim lobjTransformation As Transformation = DirectCast(value, Transformation)

      With writer

        If Helper.IsRunningInstalled Then
          .Formatting = Formatting.None
        Else
          .Formatting = Formatting.Indented
        End If
        .WriteComment("Json serialization / deserialization for transformations is in an experimental stage and should not yet be relied upon.")
        .WriteStartObject()

        ' Write the Operation Type
        .WritePropertyName("transformation")
        .WriteStartObject()

        ' Write the 'Name' property
        .WritePropertyName("name")
        .WriteValue(lobjTransformation.Name)

        ' Write the 'Description' property
        .WritePropertyName("description")
        .WriteValue(lobjTransformation.Description)

        '' Write the 'LogResult' property
        '.WritePropertyName("logresult")
        '.WriteValue(lobjTransformation.LogResult)

        '' Write the 'Locale' property
        '.WritePropertyName("locale")
        '.WriteValue(lobjTransformation.Locale.ToString())

        '' Write the 'Parameters' collection
        '.WritePropertyName("parameters")
        '.WriteStartArray()
        'For Each lobjParameter As Parameter In lobjTransformation.Parameters
        '  .WriteRawValue(lobjParameter.ToJson())
        'Next
        '.WriteEndArray()

        ' Write the 'Operations' collection
        .WritePropertyName("actions")
        .WriteStartArray()
        For Each lobjAction As Action In lobjTransformation.Actions
          .WriteRawValue(lobjAction.ToJson())
        Next
        .WriteEndArray()

        .WriteEndObject()
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
      Dim lstrName As String = String.Empty
      Dim lstrDescription As String = String.Empty
      Dim lblnLogResult As Boolean = False
      Dim lstrLocale As String = String.Empty
      Dim lobjActions As New Actions
      'Dim lobjParameters As New Parameters
      Dim lobjTransformation As Transformation

      While reader.Read
        Select Case reader.TokenType
          Case JsonToken.PropertyName
            lstrCurrentPropertyName = reader.Value

          Case JsonToken.String, JsonToken.Boolean, JsonToken.Date, JsonToken.Integer, JsonToken.Float
            Select Case lstrCurrentPropertyName
              Case "name"
                lstrName = reader.Value
              Case "description"
                lstrDescription = reader.Value
              'Case "logresult"
              '  lblnLogResult = reader.Value
              Case "locale"
                lstrLocale = reader.Value
              Case "parameters"
                lstrCurrentPropertyName = reader.Value
              Case "actions"
                lstrCurrentPropertyName = reader.Value
            End Select

          Case JsonToken.StartObject
            Select Case lstrCurrentPropertyName
              'Case "parameters"
              '  lobjParameters.Add(Parameter.CreateFromJsonReader(reader))
              Case "actions"
                lobjActions.Add(Action.CreateFromJsonReader(reader))
            End Select

        End Select
      End While

      lobjTransformation = New Transformation()  ''(lstrName, lstrDescription, CultureInfo.CreateSpecificCulture(lstrLocale))
      With lobjTransformation
        .Name = lstrName
        .Description = lstrDescription
        '.LogResult = lblnLogResult
        '.Parameters.AddRange(lobjParameters)
        .Actions.AddRange(lobjActions)
      End With

      Return lobjTransformation

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function CanConvert(objectType As Type) As Boolean
    Return objectType = GetType(Transformation)
  End Function

End Class
