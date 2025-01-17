'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Exceptions
  ''' <summary>
  ''' Thrown when the property type is not valid 
  ''' </summary>
  ''' <remarks></remarks>
  Public Class InvalidPropertyTypeException
    Inherits PropertyException

#Region "Class Variables"

    Private menuExpectedPropertyType As Core.PropertyType
    Private menuActualPropertyType As Core.PropertyType
    Private mobjProperty As Core.ECMProperty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ExpectedType() As Core.PropertyType
      Get
        Return menuExpectedPropertyType
      End Get
    End Property

    Public ReadOnly Property ActualPropertyType() As Core.PropertyType
      Get
        Return menuActualPropertyType
      End Get
    End Property

    Public ReadOnly Property InvalidProperty() As Core.ECMProperty
      Get
        Return mobjProperty
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="lpExpectedType">The type expected</param>
    ''' <param name="lpInvalidProperty">The invalid property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpExpectedType As Core.PropertyType, ByVal lpInvalidProperty As Core.ECMProperty)
      MyBase.New(String.Format("The property '{0}' is a {1} property, a {2} property was expected.",
                               lpInvalidProperty.Name, lpInvalidProperty.Type.ToString, lpExpectedType.ToString),
                             lpInvalidProperty)
      menuExpectedPropertyType = lpExpectedType
      menuActualPropertyType = lpInvalidProperty.Type
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="lpInvalidProperty">The invalid property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpInvalidProperty As Core.ECMProperty)
      MyBase.New(lpInvalidProperty)
      menuActualPropertyType = lpInvalidProperty.Type
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="lpInvalidProperty">The invalid property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpInvalidProperty As Core.IProperty)
      MyBase.New(lpInvalidProperty)
      menuActualPropertyType = lpInvalidProperty.Type
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="lpInvalidProperty">The invalid property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal lpInvalidProperty As Core.ECMProperty)
      MyBase.New(message, lpInvalidProperty)
      menuActualPropertyType = lpInvalidProperty.Type
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="lpExpectedType">The type expected</param>
    ''' <param name="lpActualType">The actual type</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal lpExpectedType As Core.PropertyType, ByVal lpActualType As Core.PropertyType)
      MyBase.New(message)

      menuExpectedPropertyType = lpExpectedType
      menuActualPropertyType = lpActualType

    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="lpExpectedType">The type expected</param>
    ''' <param name="lpActualType">The actual type</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal lpExpectedType As Core.PropertyType, ByVal lpActualType As Core.PropertyType, ByVal innerException As Exception)
      MyBase.New(message, innerException)

      menuExpectedPropertyType = lpExpectedType
      menuActualPropertyType = lpActualType

    End Sub

#End Region

  End Class

End Namespace