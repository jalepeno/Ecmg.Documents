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

  <XmlInclude(GetType(MultiValueBinaryProperty)),
  XmlInclude(GetType(MultiValueBooleanProperty)),
  XmlInclude(GetType(MultiValueBooleanProperty)),
  XmlInclude(GetType(MultiValueDateTimeProperty)),
  XmlInclude(GetType(MultiValueDoubleProperty)),
  XmlInclude(GetType(MultiValueGuidProperty)),
  XmlInclude(GetType(MultiValueHtmlProperty)),
  XmlInclude(GetType(MultiValueLongProperty)),
  XmlInclude(GetType(MultiValueObjectProperty)),
  XmlInclude(GetType(MultiValueStringProperty)),
  XmlInclude(GetType(MultiValueUriProperty)),
  XmlInclude(GetType(MultiValueXmlProperty))>
  Public MustInherit Class MultiValueProperty
    Inherits ECMProperty
    Implements IMultiValuedProperty

#Region "Public Properties"

    Public Shadows Property Values As Values Implements IMultiValuedProperty.Values
      Get
        Return MyBase.Values
      End Get
      Set(ByVal value As Values)
        MyBase.Values = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(Nothing)
      Try
        Cardinality = Core.Cardinality.ecmMultiValued
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpType As PropertyType)
      MyBase.New(Nothing)
      Try
        Cardinality = Core.Cardinality.ecmMultiValued
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
        Cardinality = Core.Cardinality.ecmMultiValued
        Type = lpType
        Name = lpName
        SystemName = lpName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpType As PropertyType, ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValues As Values)
      MyBase.New(Nothing)
      Try
        Cardinality = Core.Cardinality.ecmMultiValued
        Type = lpType
        Name = lpName
        SystemName = lpSystemName
        Values = lpValues
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Function ToSingleton(ByVal lpValueDelimiter As String) As SingletonProperty
      Try

        ' Create the new singleton property using the same type and name as me.
        Dim lobjSingleValuedProperty As SingletonProperty = PropertyFactory.Create(Type, Name, Core.Cardinality.ecmSingleValued)

        If Values.Count > 0 Then

          Select Case Type
            Case PropertyType.ecmString
              ' Delimit the values as assign them to the new value
              lobjSingleValuedProperty.Value = Values.ToDelimitedString(lpValueDelimiter)
            Case Else
              ' Since values are stringly typed, returning a delimited list 
              ' only makes sense for strings.
              ' Return the first value instead
              lobjSingleValuedProperty.Value = Values.GetFirstValue.Value
          End Select
        End If

        Return lobjSingleValuedProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace