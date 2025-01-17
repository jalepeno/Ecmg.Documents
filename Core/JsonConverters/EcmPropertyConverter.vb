'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  EcmPropertyConverter.vb
'   Description :  [type_description_here]
'   Created     :  1/15/2014 12:59:41 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.TypeConverters
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core

  Public Class EcmPropertyConverter
    Inherits JsonConverter

    Public Overrides Function CanConvert(objectType As Type) As Boolean
      Try
        If ((TypeOf objectType Is IProperty) AndAlso (TypeOf objectType IsNot IParameter)) Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function ReadJson(reader As JsonReader, objectType As Type, existingValue As Object, serializer As JsonSerializer) As Object
      Try

        Dim lstrCurrentPropertyName As String = String.Empty
        Dim lstrName As String = String.Empty
        Dim lstrSystemName As String = String.Empty
        Dim lstrDisplayName As String = String.Empty
        Dim lenuType As PropertyType
        Dim lenuCardinality As Cardinality
        Dim lobjValue As Object = Nothing
        Dim lobjValues As New Values

        While reader.Read
          Select Case reader.TokenType
            Case JsonToken.PropertyName
              lstrCurrentPropertyName = reader.Value

            Case JsonToken.String, JsonToken.Boolean, JsonToken.Date, JsonToken.Integer, JsonToken.Float

              Select Case lstrCurrentPropertyName
                Case "name"
                  lstrName = reader.Value

                Case "systemName"
                  lstrSystemName = reader.Value

                Case "displayName"
                  lstrDisplayName = reader.Value

                Case "type"
                  lenuType = EnumDescriptionTypeConverter.GetValue(GetType(PropertyType), reader.Value)

                Case "cardinality"
                  lenuCardinality = EnumDescriptionTypeConverter.GetValue(GetType(Cardinality), reader.Value)

                Case "value"
                  lobjValue = reader.Value

                Case "values"
                  lobjValues.Add(reader.Value)

              End Select

            Case JsonToken.EndObject
              Exit While

          End Select
        End While

        If String.IsNullOrEmpty(lstrSystemName) Then
          lstrSystemName = lstrName
        End If

        Dim lobjReturnProperty As IProperty = Nothing

        Select Case lenuCardinality
          Case Cardinality.ecmSingleValued
            lobjReturnProperty = PropertyFactory.Create(lenuType, lstrName, lstrSystemName, lenuCardinality, lobjValue)
          Case Cardinality.ecmMultiValued
            lobjReturnProperty = PropertyFactory.Create(lenuType, lstrName, lstrSystemName, lenuCardinality, lobjValues)
        End Select

        If String.IsNullOrEmpty(lstrDisplayName) Then
          lobjReturnProperty.DisplayName = lstrCurrentPropertyName
        End If

        Return lobjReturnProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
      Try

        Dim lobjItem As IProperty = value

        With writer
          If Helper.IsRunningInstalled Then
            .Formatting = Formatting.None
          Else
            .Formatting = Formatting.Indented
          End If

          .WriteStartObject()
          '.WritePropertyName("class")
          '.WriteValue(lobjProperty.GetType.Name)
          .WritePropertyName("name")
          .WriteValue(lobjItem.Name)

          If Not String.Equals(lobjItem.Name, lobjItem.SystemName) Then
            .WritePropertyName("systemName")
            .WriteValue(lobjItem.SystemName)
          End If

          If Not String.Equals(lobjItem.Name, lobjItem.DisplayName) Then
            .WritePropertyName("displayName")
            .WriteValue(lobjItem.DisplayName)
          End If

          If Helper.IsAssignableFrom("IParameter", value) Then
            .WritePropertyName("description")
            .WriteValue(lobjItem.Description)
          End If

          .WritePropertyName("type")
          .WriteValue(EnumDescriptionTypeConverter.GetDescription(lobjItem.Type.GetType, lobjItem.Type.ToString))
          .WritePropertyName("cardinality")
          .WriteValue(EnumDescriptionTypeConverter.GetDescription(lobjItem.Cardinality.GetType, lobjItem.Cardinality.ToString))
          '.WriteValue(lobjProperty.Cardinality.ToString)
          Select Case lobjItem.Cardinality
            Case Cardinality.ecmSingleValued
              .WritePropertyName("value")
              Select Case lobjItem.Type
                Case PropertyType.ecmString
                  .WriteValue(lobjItem.Value.ToString)

                Case PropertyType.ecmBinary
                  .WriteValue(lobjItem.Value.ToString())

                Case PropertyType.ecmBoolean
                  .WriteValue(lobjItem.Value)

                Case PropertyType.ecmDate
                  .WriteValue(lobjItem.Value)

                Case PropertyType.ecmDouble
                  .WriteValue(lobjItem.Value)

                Case PropertyType.ecmEnum
                  .WriteValue(lobjItem.Value.ToString)

                Case PropertyType.ecmGuid
                  .WriteValue(lobjItem.Value.ToString)

                Case PropertyType.ecmHtml
                  .WriteValue(lobjItem.Value.ToString)

                Case PropertyType.ecmLong
                  .WriteValue(lobjItem.Value)

                Case PropertyType.ecmUri
                  .WriteValue(lobjItem.Value.ToString)

                Case PropertyType.ecmXml
                  .WriteValue(lobjItem.Value.ToString)

                Case PropertyType.ecmUndefined
                  .WriteValue(lobjItem.Value.ToString)

              End Select

            Case Cardinality.ecmMultiValued
              .WritePropertyName("values")
              .WriteStartArray()
              For Each lobjValue As Object In lobjItem.Values
                If TypeOf lobjValue Is Value Then
                  .WriteValue(DirectCast(lobjValue, Value).ToString)
                Else
                  .WriteValue(lobjValue.ToString)
                End If
              Next
              .WriteEndArray()
          End Select
          If lobjItem.HasStandardValues AndAlso Not Helper.IsAssignableFrom("IParameter", value) Then
            .WritePropertyName("standardValues")
            .WriteStartArray()
            For Each lobjValue As Object In lobjItem.StandardValues
              .WriteValue(lobjValue.ToString)
            Next
            .WriteEndArray()
          End If
          .WriteEndObject()
        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace