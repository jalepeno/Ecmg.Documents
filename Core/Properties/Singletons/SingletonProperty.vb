'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  <XmlInclude(GetType(SingletonBinaryProperty)),
  XmlInclude(GetType(SingletonBooleanProperty)),
  XmlInclude(GetType(SingletonBooleanProperty)),
  XmlInclude(GetType(SingletonDateTimeProperty)),
  XmlInclude(GetType(SingletonDoubleProperty)),
  XmlInclude(GetType(SingletonGuidProperty)),
  XmlInclude(GetType(SingletonHtmlProperty)),
  XmlInclude(GetType(SingletonLongProperty)),
  XmlInclude(GetType(SingletonObjectProperty)),
  XmlInclude(GetType(SingletonStringProperty)),
  XmlInclude(GetType(SingletonUriProperty))>
  Public MustInherit Class SingletonProperty
    Inherits ECMProperty
    Implements ISingletonProperty

#Region "Public Properties"

    Public Shadows Property Value As Object Implements ISingletonProperty.Value
      Get
        Return MyBase.Value
      End Get
      Set(ByVal value As Object)
        MyBase.Value = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(Nothing)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpType As PropertyType)
      MyBase.New(Nothing)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
        Type = lpType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpType As PropertyType, ByVal lpName As String)
      MyBase.New(Nothing)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
        Type = lpType
        Name = lpName
        SystemName = lpName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpType As PropertyType, ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As Object)
      MyBase.New(Nothing)
      Try
        Cardinality = Core.Cardinality.ecmSingleValued
        Type = lpType
        Name = lpName
        SystemName = lpSystemName
        Value = lpValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Function ToMultiValue(lpDelimiter As String) As MultiValueProperty
      Try

        ' Create the new multivalued property using the same type and name as me.
        Dim lobjMultiValueProperty As MultiValueProperty = PropertyFactory.Create(Type, Name, Core.Cardinality.ecmMultiValued)
        Dim lobjValue As Value = Nothing

        ' Make sure we have a value
        If Value IsNot Nothing Then
          If TypeOf (Value) Is String Then
            If Not String.IsNullOrEmpty(Value) Then
              ' Add the value to the multivalued list
              Dim lstrValues() As String = Value.Split(lpDelimiter)
              For Each lstrValue As String In lstrValues
                If lstrValue.Length > 0 Then
                  lobjValue = New Value(lstrValue.Trim())
                  lobjMultiValueProperty.Values.Add(lobjValue)
                End If
              Next
            End If
          Else
            ' Add the value to the multivalued list
            lobjMultiValueProperty.Values.Add(Value)
          End If
        End If

        Return lobjMultiValueProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace