'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ClassificationLongProperty
    Inherits ClassificationProperty

#Region "Class Variables"

    Private _minValue As Nullable(Of Long)
    Private _maxValue As Nullable(Of Long)

#End Region

#Region "Attributes specific to Integer Object-Type Property Definitions"

    ' The following Object attributes MUST only apply to Property-Type 
    ' definitions whose propertyType is “Integer”, in addition to the 
    ' common attributes specified above.  A repository MAY provide 
    ' additional guidance on what values can be accepted.  If the 
    ' following attributes are not present the repository behavior 
    ' is undefined and it MAY throw an exception if a runtime constraint 
    ' is encountered.

    ''' <summary>
    ''' The minimum value allowed for this property.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If an application tries to set the value of this property 
    ''' to a value lower than minValue, the repository MUST throw 
    ''' a constraint exception.
    ''' </remarks>
    ''' <exception cref="ConstraintException">
    ''' If an application tries to set the value of 
    ''' this property to a value lower than minValue, 
    '''  a constraint exception will be thrown.
    ''' </exception>
    Public Property MinValue As Nullable(Of Long)
      Get
        If Type <> PropertyType.ecmLong Then
          Throw New ConstraintException(String.Format(
            "MinValue not valid for {0} property type, use this property only for Long property types.",
            Type.ToString))
        End If
        Return _minValue
      End Get
      Set(ByVal value As Nullable(Of Long))
        If Type <> PropertyType.ecmLong Then
          Throw New ConstraintException(String.Format(
            "MinValue not valid for {0} property type, use this property only for Long property types.",
            Type.ToString))
        End If
        _minValue = value
      End Set
    End Property

    ''' <summary>
    ''' The maximum value allowed for this property.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If an application tries to set the value of 
    ''' this property to a value higher than maxValue, 
    ''' the repository MUST throw a constraint exception.
    ''' </remarks>
    ''' <exception cref="ConstraintException">
    ''' If an application tries to set the value of 
    ''' this property to a value higher than maxValue, 
    '''  a constraint exception will be thrown.
    ''' </exception>
    Public Property MaxValue As Nullable(Of Long)
      Get
        If Type <> PropertyType.ecmLong Then
          Throw New ConstraintException(String.Format(
            "MaxValue not valid for {0} property type, use this property only for Long property types.",
            Type.ToString))
        End If
        Return _maxValue
      End Get
      Set(ByVal value As Nullable(Of Long))
        If Type <> PropertyType.ecmLong Then
          Throw New ConstraintException(String.Format(
            "MaxValue not valid for {0} property type, use this property only for Long property types.",
            Type.ToString))
        End If
        _maxValue = value
      End Set
    End Property

    Public Overloads Property DefaultValue As Nullable(Of Long)

#End Region

#Region "Constructors"

    Friend Sub New()
      MyBase.New(PropertyType.ecmLong)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmLong, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpDefaultValue As Nullable(Of Long))

      MyBase.New(PropertyType.ecmLong, lpName, lpSystemName)

      Try
        DefaultValue = lpDefaultValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Friend Sub New(ByVal lpName As String,
                   ByVal lpSystemName As String,
                   ByVal lpDefaultValue As Nullable(Of Long),
                   ByVal lpMinValue As Nullable(Of Long),
                   ByVal lpMaxValue As Nullable(Of Long))
      MyBase.New(PropertyType.ecmLong, lpName, lpSystemName)
      Try

        DefaultValue = lpDefaultValue
        MinValue = lpMinValue
        MaxValue = lpMaxValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace

