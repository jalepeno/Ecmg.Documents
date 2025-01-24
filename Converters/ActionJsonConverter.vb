' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ActionJsonConverter.vb
'  Description :  [type_description_here]
'  Created     :  01/21/2025 1:08:13 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------


#Region "Imports"

Imports Documents.Core
Imports Documents.Transformations
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Public Class ActionJsonConverter
  Inherits JsonConverter

  Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
    Try
      Dim lobjAction As IActionable = DirectCast(value, IActionable)

      With writer
        If Helper.IsRunningInstalled Then
          .Formatting = Formatting.None
        Else
          .Formatting = Formatting.Indented
        End If

        .WriteStartObject()

        ' Write the Action Type
        .WritePropertyName("type")
        .WriteValue(lobjAction.GetType.Name)

        ' Write the 'Name' property
        .WritePropertyName("name")
        .WriteValue(lobjAction.Name)

        ' Write the 'Description' property
        .WritePropertyName("description")
        .WriteValue(lobjAction.Description)

        '' Write the 'LogResult' property
        '.WritePropertyName("logresult")
        '.WriteValue(lobjAction.LogResult)

        '' Write the 'Scope' property
        '.WritePropertyName("scope")
        '.WriteValue([Enum].GetName(GetType(OperationScope), lobjAction.Scope))

        ' Write the 'Parameters' collection
        .WritePropertyName("parameters")
        .WriteStartArray()
        For Each lobjParameter As Parameter In lobjAction.Parameters
          .WriteRawValue(lobjParameter.ToJson())
        Next

        .WriteEndArray()

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
      Dim lstrDescription As String = String.Empty
      Dim lblnLogResult As Boolean = False
      Dim lstrScope As String = String.Empty
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
              Case "description"
                lstrDescription = reader.Value
              'Case "logresult"
              '  lblnLogResult = reader.Value
              'Case "scope"
              '  lstrScope = reader.Value
              Case "parameters"
                lstrCurrentPropertyName = reader.Value
                ' TODO: Handle the parameters

            End Select

          Case JsonToken.StartArray
            Select Case lstrCurrentPropertyName
              Case "parameters"
                'Beep()
                ' TODO: Handle the parameters
                lobjParameters.Add(Parameter.CreateFromJsonReader(reader))
            End Select

          Case JsonToken.StartObject
            Select Case lstrCurrentPropertyName
              Case "parameters"
                'Beep()
                ' TODO: Handle the parameters
                lobjParameters.Add(Parameter.CreateFromJsonReader(reader))
            End Select
          Case JsonToken.EndArray
            Dim lstrActionName As String = lstrType.Replace("Action", String.Empty)
            If ActionFactory.Instance.AvailableActions.ContainsKey(lstrActionName) Then
              Dim lobjActionType As Type = ActionFactory.Instance.AvailableActions.Item(lstrActionName)
              Dim lobjAction As ITransformationAction = ActionFactory.Create(lstrActionName)
              With lobjAction
                '.LogResult = lblnLogResult
                '.Scope = [Enum].Parse(GetType(OperationScope), lstrScope)
                .Parameters = lobjParameters
              End With
              Return lobjAction
            End If
        End Select
      End While

      Return Nothing

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function CanConvert(objectType As Type) As Boolean
    If TypeOf objectType Is IActionable OrElse Helper.IsAssignableFrom("IActionable", objectType) Then
      Return True
    Else
      Return False
    End If
  End Function
End Class
