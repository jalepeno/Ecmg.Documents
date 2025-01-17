'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ParameterFactory.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 2:19:07 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Exceptions
Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ParameterFactory

#Region "Public Methods"

#Region "Create Implementations"

#Region "Overloaded Create Signatures"

    Public Shared Function Create(ByVal lpProperty As Parameter) As IParameter
      Try
        Dim lobjReturnProperty As Parameter
        If Not String.IsNullOrEmpty(lpProperty.SystemName) Then
          lobjReturnProperty = Create(lpProperty.Type, lpProperty.Name, lpProperty.SystemName, lpProperty.Cardinality)
        Else
          lobjReturnProperty = Create(lpProperty.Type, lpProperty.Name, lpProperty.PackedName, lpProperty.Cardinality)
        End If

        Select Case lpProperty.Cardinality
          Case Cardinality.ecmSingleValued
            lobjReturnProperty.Value = lpProperty.Value
          Case Else
            For Each lobjValue As Object In lpProperty.Values
              lobjReturnProperty.Values.Add(lobjValue)
            Next
        End Select

        If TypeOf lpProperty Is SingletonEnumParameter Then
          CType(lobjReturnProperty, SingletonEnumParameter).SetEnumType(CType(lpProperty, SingletonEnumParameter).EnumType)
        End If

        Return lobjReturnProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpProperty As IParameter) As IParameter
      Try
        Dim lobjReturnProperty As IParameter
        If Not String.IsNullOrEmpty(lpProperty.SystemName) Then
          lobjReturnProperty = Create(lpProperty.Type, lpProperty.Name, lpProperty.SystemName, lpProperty.Cardinality)
        Else
          lobjReturnProperty = Create(lpProperty.Type, lpProperty.Name, lpProperty.Name, lpProperty.Cardinality)
        End If

        lobjReturnProperty.Description = lpProperty.Description

        If lpProperty.Type = PropertyType.ecmEnum Then
          CType(lobjReturnProperty, Parameter).EnumTypeName = CType(lpProperty, Parameter).EnumTypeName
        End If

        Select Case lpProperty.Cardinality
          Case Cardinality.ecmSingleValued
            lobjReturnProperty.Value = lpProperty.Value
          Case Else
            For Each lobjValue As Object In lpProperty.Values
              lobjReturnProperty.Values.Add(lobjValue)
            Next
        End Select

        Return lobjReturnProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
            ByVal lpName As String,
            ByVal lpValue As String,
            ByVal lpDescription As String) As IParameter
      Try
        Dim lobjParameter As IParameter = Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
        lobjParameter.Description = lpDescription
        Return lobjParameter
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
        ByVal lpName As String,
        ByVal lpValue As Object,
        ByVal lpDescription As String) As IParameter
      Try
        Dim lobjParameter As IParameter = Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
        lobjParameter.Description = lpDescription
        Return lobjParameter
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
        ByVal lpName As String,
        ByVal lpValue As Nullable(Of Boolean),
        ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue, True, Nothing, lpDescription)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
            ByVal lpName As String,
            ByVal lpValue As Nullable(Of DateTime),
            ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As Single,
      ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, New Nullable(Of Single)(lpValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
        ByVal lpName As String,
        ByVal lpValue As Nullable(Of Single),
        ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As Double,
      ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, New Nullable(Of Double)(lpValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
        ByVal lpName As String,
        ByVal lpValue As Nullable(Of Double),
        ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As Integer,
      ByVal lpDescription As String) As IParameter
      Try
        Dim lobjParameter As IParameter
        If lpType = PropertyType.ecmEnum Then
          lobjParameter = Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
        Else
          lobjParameter = Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, New Nullable(Of Integer)(lpValue))
        End If

        lobjParameter.Description = lpDescription

        Return lobjParameter

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As [Enum],
      ByVal lpEnumType As Type,
      ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue, True, lpEnumType, lpDescription)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As String,
      ByVal lpEnumType As Type,
      ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue, True, lpEnumType, lpDescription)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As Nullable(Of Integer),
      ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As Long,
      ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, New Nullable(Of Long)(lpValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
        ByVal lpName As String,
        ByVal lpValue As Nullable(Of Long),
        ByVal lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, lpValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
            ByVal lpName As String,
            ByVal lpCardinality As Cardinality) As IParameter
      Try
        Return Create(lpType, lpName, lpName, lpCardinality)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
                  ByVal lpName As String,
                  ByVal lpSystemName As String,
                  ByVal lpCardinality As Cardinality) As IParameter
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
                  ByVal lpValue As Object) As IParameter
      Try
        Return Create(lpType, lpName, lpSystemName, lpCardinality, lpValue, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
                                  ByVal lpName As String,
                                  lpValue As MapList,
                                  lpDescription As String) As IParameter
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmMultiValued, lpValue, True, Nothing, lpDescription)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Create Implementation"

    Public Shared Function Create(ByVal lpType As PropertyType,
          ByVal lpName As String,
          ByVal lpSystemName As String,
          ByVal lpCardinality As Cardinality,
          ByVal lpValue As Object,
          ByVal lpPersistent As Boolean) As IParameter
      Try
        Return Create(lpType, lpName, lpSystemName, lpCardinality, lpValue, lpPersistent, Nothing)
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
          ByVal lpValue As Object,
          ByVal lpPersistent As Boolean,
          ByVal lpEnumType As Type) As IParameter
      Try
        Return Create(lpType, lpName, lpSystemName, lpCardinality, lpValue, lpPersistent, lpEnumType, String.Empty)
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
              ByVal lpValue As Object,
              ByVal lpPersistent As Boolean,
              ByVal lpEnumType As Type,
              ByVal lpDescription As String) As IParameter
      Try

        Dim lobjReturnValue As Parameter = Nothing

        Select Case lpCardinality
          Case Core.Cardinality.ecmSingleValued
            Select Case lpType

              Case PropertyType.ecmString
                lobjReturnValue = New SingletonStringParameter(lpName, lpSystemName, lpValue)

              Case PropertyType.ecmBoolean
                If lpValue IsNot Nothing Then
                  If TypeOf (lpValue) Is String Then
                    If Not String.IsNullOrEmpty(lpValue) Then
                      lobjReturnValue = New SingletonBooleanParameter(lpName, lpSystemName, lpValue.ToString)
                    Else
                      lobjReturnValue = New SingletonBooleanParameter(lpName, lpSystemName)
                    End If
                  Else
                    lobjReturnValue = New SingletonBooleanParameter(lpName, lpSystemName, lpValue.ToString)
                  End If
                Else
                  lobjReturnValue = New SingletonBooleanParameter(lpName, lpSystemName)
                End If

              Case PropertyType.ecmDate
                If lpValue IsNot Nothing Then
                  If TypeOf (lpValue) Is DateTime Then
                    lobjReturnValue = New SingletonDateTimeParameter(lpName, lpSystemName, lpValue)
                  ElseIf TypeOf (lpValue) Is String Then
                    If Not String.IsNullOrEmpty(lpValue) Then
                      lobjReturnValue = New SingletonDateTimeParameter(lpName, lpSystemName, CDate(lpValue))
                    Else
                      lobjReturnValue = New SingletonDateTimeParameter(lpName, lpSystemName, Nothing)
                    End If
                  Else
                    Throw New ArgumentException(String.Format("The value '{0}' supplied for property '{1}' is not a valid date.", lpValue.ToString, lpName))
                  End If
                Else
                  lobjReturnValue = New SingletonDateTimeParameter(lpName)
                  lobjReturnValue.SystemName = lpSystemName
                End If

              Case PropertyType.ecmLong
                If TypeOf lpValue Is [Enum] Then
                  lobjReturnValue = New SingletonEnumParameter(lpName, lpSystemName, lpValue)
                ElseIf Not TypeOf (lpValue) Is String Then
                  lobjReturnValue = New SingletonLongParameter(lpName, lpSystemName, New Nullable(Of Long)(lpValue))
                Else
                  If Not String.IsNullOrEmpty(lpValue) Then
                    lobjReturnValue = New SingletonLongParameter(lpName, lpSystemName, New Nullable(Of Long)(lpValue))
                  Else
                    lobjReturnValue = New SingletonLongParameter(lpName, lpSystemName, Nothing)
                  End If
                End If

              Case PropertyType.ecmEnum
                If TypeOf lpValue Is String Then
                  lobjReturnValue = New SingletonEnumParameter(lpName, lpSystemName, [Enum].Parse(lpEnumType, lpValue), lpEnumType)
                Else
                  lobjReturnValue = New SingletonEnumParameter(lpName, lpSystemName, lpValue, lpEnumType)
                End If

              Case PropertyType.ecmDouble
                If Not TypeOf (lpValue) Is String Then
                  lobjReturnValue = New SingletonDoubleParameter(lpName, lpSystemName, New Nullable(Of Double)(lpValue))
                Else
                  If Not String.IsNullOrEmpty(lpValue) Then
                    lobjReturnValue = New SingletonDoubleParameter(lpName, lpSystemName, New Nullable(Of Double)(lpValue))
                  Else
                    lobjReturnValue = New SingletonDoubleParameter(lpName, lpSystemName, Nothing)
                  End If
                End If

              Case PropertyType.ecmObject
                lobjReturnValue = New SingletonObjectParameter(lpName, lpSystemName, lpValue)

              Case PropertyType.ecmUndefined
                lobjReturnValue = New SingletonStringParameter(lpName, lpSystemName, lpValue)

              Case PropertyType.ecmValueMap
                Throw New ArgumentOutOfRangeException("lpType", "Invalid combination: Singlevalued ValueMap")
              Case Else
                lobjReturnValue = New SingletonStringParameter(lpName, lpSystemName, lpValue)

            End Select
          Case Core.Cardinality.ecmMultiValued
            ' Initialize the values
            Dim lobjValues As Values
            Dim lobjCleanValues As New Values
            If TypeOf lpValue Is Values Then
              lobjValues = lpValue
              ' Check to see if any or all of the internal values are using the old Value object
              For Each lobjValue As Object In lobjValues
                If TypeOf lobjValue Is Value Then
                  ' If so then add the actual data value to a clean collection.
                  lobjCleanValues.Add(DirectCast(lobjValue, Value).Value)
                Else
                  ' Otherwise just add the value as is.
                  lobjCleanValues.Add(lobjValue)
                End If
              Next
              ' Set the values collection to the clean values collection.
              lobjValues = lobjCleanValues
            Else
              lobjValues = New Values
              If TypeOf lpValue Is String Then
                If String.IsNullOrEmpty(lpValue) = False Then
                  lobjValues.Add(lpValue)
                End If
#If SILVERLIGHT <> 1 Then
              ElseIf TypeOf lpValue Is IEnumerable Then
                For Each lobjValue As Object In lpValue
                  lobjValues.Add(lobjValue)
                Next
#End If
              Else
                If lpValue IsNot Nothing Then
                  lobjValues.Add(lpValue)
                End If
              End If
            End If

            Select Case lpType
              Case PropertyType.ecmString
                lobjReturnValue = New MultiValueStringParameter(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmBoolean
                '  lobjReturnValue = New MultiValueBooleanProperty(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmDate
                '  lobjReturnValue = New MultiValueDateTimeProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmLong
                lobjReturnValue = New MultiValueLongParameter(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmDouble
                lobjReturnValue = New MultiValueDoubleParameter(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmObject
                '  lobjReturnValue = New MultiValueObjectProperty(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmBinary
                '  lobjReturnValue = New MultiValueBinaryProperty(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmGuid
                '  lobjReturnValue = New MultiValueGuidProperty(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmHtml
                '  lobjReturnValue = New MultiValueHtmlProperty(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmUri
                '  lobjReturnValue = New MultiValueUriProperty(lpName, lpSystemName, lobjValues)
                'Case PropertyType.ecmXml
                '  lobjReturnValue = New MultiValueXmlProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmValueMap
                lobjReturnValue = New MapListParameter(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmUndefined
                lobjReturnValue = New MultiValueStringParameter(lpName, lpSystemName, lobjValues)
              Case Else
                lobjReturnValue = New MultiValueStringParameter(lpName, lpSystemName, lobjValues)

            End Select
        End Select

        ' Set the persistence
        lobjReturnValue.SetPersistence(lpPersistent)

        ' ' Set the description
        lobjReturnValue.Description = lpDescription

        Return lobjReturnValue

      Catch InvCastEx As InvalidCastException
        ApplicationLogging.LogException(InvCastEx, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry(String.Format("Unable to cast value '{0}' to type {1} for property '{2}', the specified cast is invalid.",
                                                       lpValue.ToString, lpType.ToString, lpName), TraceEventType.Error, 62931)
        ' Re-throw the exception to the caller
        Throw

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#End Region

    Public Shared Function ConvertStringToEnum(lpOriginalParameter As SingletonStringParameter, lpEnumType As Type) As IParameter
      Try
        Dim lobjValue As [Enum] = [Enum].Parse(lpEnumType, lpOriginalParameter.Value)

        Return Create(PropertyType.ecmEnum, lpOriginalParameter.Name, lobjValue, lpEnumType, lpOriginalParameter.Description)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Sub RenameParameter(ByRef lpParameter As IParameter, lpNewName As String)
      Try
        lpParameter.Name = lpNewName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub UpdateParameterToEnum(ByRef lpParameters As IParameters, lpName As String, lpEnumType As Type)
      Try
        If lpParameters Is Nothing Then
          Throw New ArgumentNullException("lpParameters")
        End If

        Dim lobjOriginalParameter As IParameter = lpParameters.GetItemByName(lpName)
        If lobjOriginalParameter Is Nothing Then
          Throw New ItemDoesNotExistException(lpName)
        End If

        If lobjOriginalParameter.Type = PropertyType.ecmString Then
          Dim lobjReplacementParameter As SingletonEnumParameter = ParameterFactory.ConvertStringToEnum(lobjOriginalParameter, lpEnumType)
          lpParameters.Remove(lobjOriginalParameter)
          lpParameters.Add(lobjReplacementParameter)
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