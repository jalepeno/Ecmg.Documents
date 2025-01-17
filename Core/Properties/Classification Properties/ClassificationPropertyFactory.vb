'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ClassificationPropertyFactory

#Region "Public Methods"

    Public Shared Function Create(ByVal lpType As PropertyType,
                                  ByVal lpCardinality As Cardinality) As IProperty
      Try

        Dim lobjReturnValue As ClassificationProperty = Nothing

        Select Case lpType

          Case PropertyType.ecmString, PropertyType.ecmUndefined
            lobjReturnValue = New ClassificationStringProperty

          Case PropertyType.ecmBinary
            lobjReturnValue = New ClassificationBinaryProperty

          Case PropertyType.ecmBoolean
            lobjReturnValue = New ClassificationBooleanProperty

          Case PropertyType.ecmDate
            lobjReturnValue = New ClassificationDateTimeProperty

          Case PropertyType.ecmDouble
            lobjReturnValue = New ClassificationDoubleProperty

          Case PropertyType.ecmGuid
            lobjReturnValue = New ClassificationGuidProperty

          Case PropertyType.ecmHtml
            lobjReturnValue = New ClassificationHtmlProperty

          Case PropertyType.ecmLong
            lobjReturnValue = New ClassificationLongProperty

          Case PropertyType.ecmObject
            lobjReturnValue = New ClassificationObjectProperty

          Case PropertyType.ecmUri
            lobjReturnValue = New ClassificationUriProperty

          Case Else
            lobjReturnValue = New ClassificationStringProperty

        End Select

        lobjReturnValue.Cardinality = lpCardinality

        Return lobjReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
                                  ByVal lpName As String,
                                  ByVal lpCardinality As Cardinality) As IProperty
      Try
        Return Create(lpType, lpName, lpName, lpCardinality, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
                              ByVal lpName As String,
                              ByVal lpSystemName As String,
                              ByVal lpCardinality As Cardinality) As IProperty
      Try
        Return Create(lpType, lpName, lpSystemName, lpCardinality, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
                                  ByVal lpName As String,
                                  ByVal lpSystemName As String,
                                  ByVal lpCardinality As Cardinality,
                                  ByVal lpDefaultValue As Object) As IProperty
      Try

        Dim lobjReturnValue As ClassificationProperty = Nothing

        If String.IsNullOrEmpty(lpName) AndAlso Not String.IsNullOrEmpty(lpSystemName) Then
          lpName = lpSystemName
        End If

        Select Case lpType

          Case PropertyType.ecmString, PropertyType.ecmUndefined
            If lpDefaultValue Is Nothing Then
              lobjReturnValue = New ClassificationStringProperty(lpName, lpSystemName, Nothing)
            Else
              lobjReturnValue = New ClassificationStringProperty(lpName, lpSystemName, lpDefaultValue.ToString)
            End If

          Case PropertyType.ecmBinary
            lobjReturnValue = New ClassificationBinaryProperty(lpName, lpSystemName)

          Case PropertyType.ecmBoolean
            'Dim lblnDefaultValue As Nullable(Of Boolean)
            Dim lblnDefaultValue As Boolean
            If lpDefaultValue IsNot Nothing AndAlso Boolean.TryParse(lpDefaultValue, lblnDefaultValue) Then
              lobjReturnValue = New ClassificationBooleanProperty(lpName, lpSystemName, lblnDefaultValue)
            Else
              lobjReturnValue = New ClassificationBooleanProperty(lpName, lpSystemName, Nothing)
            End If

          Case PropertyType.ecmDate
            'Dim ldatDefaultValue As Nullable(Of DateTime) = DateTime.MinValue
            Dim ldatDefaultValue As DateTime = DateTime.MinValue
            If lpDefaultValue IsNot Nothing AndAlso DateTime.TryParse(lpDefaultValue, ldatDefaultValue) Then
              lobjReturnValue = New ClassificationDateTimeProperty(lpName, lpSystemName, ldatDefaultValue)
            Else
              lobjReturnValue = New ClassificationDateTimeProperty(lpName, lpSystemName, Nothing)
            End If

          Case PropertyType.ecmDouble
            'Dim ldblDefaultValue As Nullable(Of Double)
            Dim ldblDefaultValue As Double
            If lpDefaultValue IsNot Nothing AndAlso Double.TryParse(lpDefaultValue, ldblDefaultValue) Then
              lobjReturnValue = New ClassificationDoubleProperty(lpName, lpSystemName, ldblDefaultValue)
            Else
              lobjReturnValue = New ClassificationDoubleProperty(lpName, lpSystemName, Nothing)
            End If

          Case PropertyType.ecmGuid

            If lpDefaultValue IsNot Nothing Then
              Try
                Dim lobjDefault As New Guid(lpDefaultValue.ToString)
                lobjReturnValue = New ClassificationGuidProperty(lpName, lpSystemName, lobjDefault)
              Catch ex As Exception
                lobjReturnValue = New ClassificationGuidProperty(lpName, lpSystemName)
              End Try
            Else
              lobjReturnValue = New ClassificationGuidProperty(lpName, lpSystemName)
            End If

          Case PropertyType.ecmHtml
            lobjReturnValue = New ClassificationHtmlProperty(lpName, lpSystemName, lpDefaultValue)

          Case PropertyType.ecmLong
            'Dim llngDefaultValue As Nullable(Of Long)
            Dim llngDefaultValue As Long
            If lpDefaultValue IsNot Nothing AndAlso Long.TryParse(lpDefaultValue, llngDefaultValue) Then
              lobjReturnValue = New ClassificationLongProperty(lpName, lpSystemName, llngDefaultValue)
            Else
              lobjReturnValue = New ClassificationLongProperty(lpName, lpSystemName, Nothing)
            End If

          Case PropertyType.ecmObject
            lobjReturnValue = New ClassificationObjectProperty(lpName, lpSystemName, lpDefaultValue)

          Case PropertyType.ecmUri
            Dim lobjDefaultValue As Uri = Nothing
            If Uri.TryCreate("", UriKind.RelativeOrAbsolute, lobjDefaultValue) = True Then
              lobjReturnValue = New ClassificationUriProperty(lpName, lpSystemName, lobjDefaultValue)
            Else
              lobjReturnValue = New ClassificationUriProperty(lpName)
              lobjReturnValue.SystemName = lpSystemName
            End If

          Case Else
            If lpDefaultValue IsNot Nothing Then
              lobjReturnValue = New ClassificationStringProperty(lpName, lpSystemName, lpDefaultValue.ToString)
            Else
              lobjReturnValue = New ClassificationStringProperty(lpName, lpSystemName, String.Empty)
            End If

        End Select

        ' Set the cardinality
        lobjReturnValue.Cardinality = lpCardinality

        Return lobjReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(lpProperty As IProperty) As IProperty
      Try
        Dim lobjNewProperty As ClassificationProperty = Create(lpProperty.Type, lpProperty.Name, lpProperty.SystemName, lpProperty.Cardinality)

        If lpProperty.HasValue Then
          If lpProperty.Cardinality = Cardinality.ecmSingleValued Then
            lobjNewProperty.Value = lpProperty.Value
          Else
            lobjNewProperty.Values = lpProperty.Values
          End If
        End If

        Return lobjNewProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
