'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


Imports System.Xml
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class PropertyFactory

#Region "Public Methods"

#Region "Create Implementations"

#Region "Overloaded Create Signatures"

    Public Shared Function Create(ByVal lpProperty As ECMProperty) As IProperty
      Try
        Dim lobjReturnProperty As ECMProperty
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

        Return lobjReturnProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpProperty As IProperty) As IProperty
      Try
        Dim lobjReturnProperty As IProperty
        If Not String.IsNullOrEmpty(lpProperty.SystemName) Then
          lobjReturnProperty = Create(lpProperty.Type, lpProperty.Name, lpProperty.SystemName, lpProperty.Cardinality)
        Else
          lobjReturnProperty = Create(lpProperty.Type, lpProperty.Name, lpProperty.Name, lpProperty.Cardinality)
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
                ByVal lpValue As String) As IProperty
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
        ByVal lpValue As Nullable(Of Boolean)) As IProperty
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
            ByVal lpValue As Nullable(Of DateTime)) As IProperty
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
      ByVal lpValue As Single) As IProperty
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
        ByVal lpValue As Nullable(Of Single)) As IProperty
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
      ByVal lpValue As Double) As IProperty
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
        ByVal lpValue As Nullable(Of Double)) As IProperty
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
      ByVal lpValue As Integer) As IProperty
      Try
        Return Create(lpType, lpName, lpName, Cardinality.ecmSingleValued, New Nullable(Of Integer)(lpValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Create(ByVal lpType As PropertyType,
      ByVal lpName As String,
      ByVal lpValue As Nullable(Of Integer)) As IProperty
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
      ByVal lpValue As Long) As IProperty
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
        ByVal lpValue As Nullable(Of Long)) As IProperty
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
            ByVal lpCardinality As Cardinality) As IProperty
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
                  ByVal lpValue As Object) As IProperty
      Try
        Return Create(lpType, lpName, lpSystemName, lpCardinality, lpValue, True)
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
                  ByVal lpPersistent As Boolean) As IProperty
      Try

        Dim lobjReturnValue As ECMProperty = Nothing

        Select Case lpCardinality
          Case Core.Cardinality.ecmSingleValued
            Select Case lpType
              Case PropertyType.ecmString
                lobjReturnValue = New SingletonStringProperty(lpName, lpSystemName, lpValue)
              Case PropertyType.ecmBoolean
                If lpValue IsNot Nothing Then
                  If TypeOf (lpValue) Is String Then
                    If Not String.IsNullOrEmpty(lpValue) Then
                      lobjReturnValue = New SingletonBooleanProperty(lpName, lpSystemName, lpValue.ToString)
                    Else
                      lobjReturnValue = New SingletonBooleanProperty(lpName, lpSystemName)
                    End If
                  Else
                    lobjReturnValue = New SingletonBooleanProperty(lpName, lpSystemName, lpValue.ToString)
                  End If
                Else
                  lobjReturnValue = New SingletonBooleanProperty(lpName, lpSystemName)
                End If
              Case PropertyType.ecmDate
                If lpValue IsNot Nothing Then
                  If TypeOf (lpValue) Is DateTime Then
                    lobjReturnValue = New SingletonDateTimeProperty(lpName, lpSystemName, lpValue)
                  ElseIf TypeOf (lpValue) Is String Then
                    If Not String.IsNullOrEmpty(lpValue) Then
                      lobjReturnValue = New SingletonDateTimeProperty(lpName, lpSystemName, CDate(lpValue))
                    Else
                      lobjReturnValue = New SingletonDateTimeProperty(lpName, lpSystemName, Nothing)
                    End If
                  Else
                    Throw New ArgumentException(String.Format("The value '{0}' supplied for property '{1}' is not a valid date.", lpValue.ToString, lpName))
                  End If
                Else
                  lobjReturnValue = New SingletonDateTimeProperty(lpName) With {
                    .SystemName = lpSystemName
                  }
                End If

              Case PropertyType.ecmLong
                If TypeOf (lpValue) IsNot String Then
                  lobjReturnValue = New SingletonLongProperty(lpName, lpSystemName, New Nullable(Of Long)(lpValue))
                Else
                  If Not String.IsNullOrEmpty(lpValue) Then
                    lobjReturnValue = New SingletonLongProperty(lpName, lpSystemName, New Nullable(Of Long)(lpValue))
                  Else
                    lobjReturnValue = New SingletonLongProperty(lpName, lpSystemName, Nothing)
                  End If
                End If

              Case PropertyType.ecmDouble
                If TypeOf (lpValue) IsNot String Then
                  lobjReturnValue = New SingletonDoubleProperty(lpName, lpSystemName, New Nullable(Of Double)(lpValue))
                Else
                  If Not String.IsNullOrEmpty(lpValue) Then
                    lobjReturnValue = New SingletonDoubleProperty(lpName, lpSystemName, New Nullable(Of Double)(lpValue))
                  Else
                    lobjReturnValue = New SingletonDoubleProperty(lpName, lpSystemName, Nothing)
                  End If
                End If

              Case PropertyType.ecmObject
                lobjReturnValue = New SingletonObjectProperty(lpName, lpSystemName, lpValue)

              Case PropertyType.ecmBinary
                If lpValue Is Nothing Then
                  lobjReturnValue = New SingletonBinaryProperty(lpName, lpSystemName, Nothing)
                ElseIf TypeOf lpValue Is Integer Then
                  lobjReturnValue = New SingletonBinaryProperty(lpName, lpSystemName, lpValue)
                ElseIf TypeOf lpValue Is String Then
                  If Not IsNumeric(lpValue) Then
                    Throw New ArgumentOutOfRangeException("lpValue", lpValue, "Numeric value expected.")
                  End If
                  Dim lintValue As Integer = CInt(lpValue)
                  Select Case lintValue
                    Case 0, 1
                      lobjReturnValue = New SingletonBinaryProperty(lpName, lpSystemName, lintValue)
                    Case Else
                      Throw New ArgumentOutOfRangeException("lpValue", lpValue, "Zero or one expected.")
                  End Select
                End If

              Case PropertyType.ecmGuid
                If TypeOf (lpValue) Is String Then
                  If Not String.IsNullOrEmpty(lpValue) Then
                    Dim lobjGuid As New Guid(lpValue.ToString)
                    lobjReturnValue = New SingletonGuidProperty(lpName, lpSystemName, lobjGuid)
                  Else
                    lobjReturnValue = New SingletonGuidProperty(lpName, lpSystemName, Nothing)
                  End If
                ElseIf TypeOf (lpValue) Is Guid Then
                  lobjReturnValue = New SingletonGuidProperty(lpName, lpSystemName, lpValue)
                ElseIf lpValue Is Nothing Then
                  lobjReturnValue = New SingletonGuidProperty(lpName, lpSystemName, Nothing)
                Else
                  Throw New ArgumentException("Guid or String expected for parameter lpValue", NameOf(lpValue))
                End If
              Case PropertyType.ecmHtml
                lobjReturnValue = New SingletonHtmlProperty(lpName, lpSystemName, lpValue)
              Case PropertyType.ecmUri
                lobjReturnValue = New SingletonUriProperty(lpName, lpSystemName, lpValue)
              Case PropertyType.ecmXml
                ' lobjReturnValue = CreateSingletonXmlProperty(lpName, lpSystemName, lpValue)
                If lpValue Is Nothing Then
                  lobjReturnValue = New SingletonXmlProperty(lpName, lpSystemName, String.Empty)
                ElseIf TypeOf lpValue Is XmlDocument Then
                  lobjReturnValue = New SingletonXmlProperty(lpName, lpSystemName, DirectCast(lpValue, XmlDocument))
                ElseIf TypeOf lpValue Is String Then
                  lobjReturnValue = New SingletonXmlProperty(lpName, lpSystemName, DirectCast(lpValue, String))
                Else
                  Throw New ArgumentOutOfRangeException(NameOf(lpValue), "Expected type is XmlDocument or String")
                End If


              Case PropertyType.ecmUndefined
                lobjReturnValue = New SingletonStringProperty(lpName, lpSystemName, lpValue)
              Case Else
                lobjReturnValue = New SingletonStringProperty(lpName, lpSystemName, lpValue)
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
                lobjReturnValue = New MultiValueStringProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmBoolean
                lobjReturnValue = New MultiValueBooleanProperty(lpName, lpSystemName, ValidateBooleanValues(lobjValues))
              Case PropertyType.ecmDate
                lobjReturnValue = New MultiValueDateTimeProperty(lpName, lpSystemName, ValidateDateValues(lobjValues))
              Case PropertyType.ecmLong
                lobjReturnValue = New MultiValueLongProperty(lpName, lpSystemName, ValidateLongValues(lobjValues))
              Case PropertyType.ecmDouble
                lobjReturnValue = New MultiValueDoubleProperty(lpName, lpSystemName, ValidateDoubleValues(lobjValues))
              Case PropertyType.ecmObject
                lobjReturnValue = New MultiValueObjectProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmBinary
                lobjReturnValue = New MultiValueBinaryProperty(lpName, lpSystemName, ValidateBinaryValues(lobjValues))
              Case PropertyType.ecmGuid
                lobjReturnValue = New MultiValueGuidProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmHtml
                lobjReturnValue = New MultiValueHtmlProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmUri
                lobjReturnValue = New MultiValueUriProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmXml
                lobjReturnValue = New MultiValueXmlProperty(lpName, lpSystemName, lobjValues)
              Case PropertyType.ecmUndefined
                lobjReturnValue = New MultiValueStringProperty(lpName, lpSystemName, lobjValues)
              Case Else
                lobjReturnValue = New MultiValueStringProperty(lpName, lpSystemName, lobjValues)
            End Select

        End Select

        ' Set the persistence
        lobjReturnValue.SetPersistence(lpPersistent)

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

#End Region

#Region "Private Methods"

    Private Shared Function ValidateBinaryValues(lpOriginalValues As Values) As Values
      Try

        Dim lobjReturnValues As New Values

        For Each lobjValue As Object In lpOriginalValues
          If TypeOf lobjValue Is Integer Then
            lobjReturnValues.Add(lobjValue)
          ElseIf TypeOf lobjValue Is String Then
            If Not IsNumeric(lobjValue) Then
              Throw New ArgumentOutOfRangeException("lpOriginalValues", lobjValue, "Numeric value expected.")
            End If
            Dim lintValue As Integer = CInt(lobjValue)
            Select Case lintValue
              Case 0, 1
                lobjReturnValues.Add(lintValue)
              Case Else
                Throw New ArgumentOutOfRangeException("lpOriginalValues", lobjValue, "Zero or one expected.")
            End Select
          End If
        Next

        Return lobjReturnValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ValidateBooleanValues(lpOriginalValues As Values) As Values
      Try

        Dim lobjReturnValues As New Values

        For Each lobjValue As Object In lpOriginalValues
          If TypeOf lobjValue Is Boolean Then
            lobjReturnValues.Add(lobjValue)
          ElseIf TypeOf lobjValue Is String Then
            Dim lblnValue As Boolean
            If Boolean.TryParse(lobjValue, lblnValue) = False Then
              Throw New ArgumentOutOfRangeException("lpOriginalValues", lobjValue, "Boolean value expected.")
            End If
            lobjReturnValues.Add(lblnValue)
          End If
        Next

        Return lobjReturnValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ValidateDateValues(lpOriginalValues As Values) As Values
      Try

        Dim lobjReturnValues As New Values

        For Each lobjValue As Object In lpOriginalValues
          If TypeOf lobjValue Is DateTime Then
            lobjReturnValues.Add(lobjValue)
          ElseIf TypeOf lobjValue Is String Then
            Dim ldatValue As DateTime
            If DateTime.TryParse(lobjValue, ldatValue) = False Then
              Throw New ArgumentOutOfRangeException("lpOriginalValues", lobjValue, "Date value expected.")
            End If
            lobjReturnValues.Add(ldatValue)
          End If
        Next

        Return lobjReturnValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ValidateDoubleValues(lpOriginalValues As Values) As Values
      Try

        Dim lobjReturnValues As New Values

        For Each lobjValue As Object In lpOriginalValues
          If TypeOf lobjValue Is Double Then
            lobjReturnValues.Add(lobjValue)
          ElseIf TypeOf lobjValue Is String Then
            Dim ldblValue As Double
            If Double.TryParse(lobjValue, ldblValue) = False Then
              Throw New ArgumentOutOfRangeException("lpOriginalValues", lobjValue, "Double value expected.")
            End If
            lobjReturnValues.Add(ldblValue)
          End If
        Next

        Return lobjReturnValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function ValidateLongValues(lpOriginalValues As Values) As Values
      Try

        Dim lobjReturnValues As New Values

        For Each lobjValue As Object In lpOriginalValues
          If TypeOf lobjValue Is Integer Then
            lobjReturnValues.Add(lobjValue)
          ElseIf TypeOf lobjValue Is Long Then
            lobjReturnValues.Add(lobjValue)
          ElseIf TypeOf lobjValue Is String Then
            Dim lintValue As Integer
            Dim ldblValue As Long
            If Integer.TryParse(CStr(lobjValue), lintValue) = True Then
              lobjReturnValues.Add(lintValue)
            ElseIf Long.TryParse(lobjValue, ldblValue) = True Then
              lobjReturnValues.Add(ldblValue)
            Else
              Throw New ArgumentOutOfRangeException("lpOriginalValues", lobjValue, "Long value expected.")
            End If

          End If
        Next

        Return lobjReturnValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
