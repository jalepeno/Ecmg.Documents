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

  Public Class ClassificationDateTimeProperty
    Inherits ClassificationProperty

#Region "Class Variables"

    Private _minValue As Nullable(Of DateTime)
    Private _maxValue As Nullable(Of DateTime)

#End Region

#Region "Attributes specific to Integer Object-Type Property Definitions"

    ' The following Object attributes MUST only apply to Property-Type 
    ' definitions whose propertyType is “DateTime”, in addition to the 
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
    Public Property MinValue As Nullable(Of DateTime)
      Get
        If Type <> PropertyType.ecmDate Then
          Throw New ConstraintException(String.Format(
            "MinValue not valid for {0} property type, use this property only for Date property types.",
            Type.ToString))
        End If
        Return _minValue
      End Get
      Set(ByVal value As Nullable(Of DateTime))
        If Type <> PropertyType.ecmDate Then
          Throw New ConstraintException(String.Format(
            "MinValue not valid for {0} property type, use this property only for Date property types.",
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
    Public Property MaxValue As Nullable(Of DateTime)
      Get
        If Type <> PropertyType.ecmDate Then
          Throw New ConstraintException(String.Format(
            "MaxValue not valid for {0} property type, use this property only for Date property types.",
            Type.ToString))
        End If
        Return _maxValue
      End Get
      Set(ByVal value As Nullable(Of DateTime))
        If Type <> PropertyType.ecmDate Then
          Throw New ConstraintException(String.Format(
            "MaxValue not valid for {0} property type, use this property only for Date property types.",
            Type.ToString))
        End If
        _maxValue = value
      End Set
    End Property

    Public Overloads Property DefaultValue As Nullable(Of DateTime)

#End Region

#Region "Constructors"

    Friend Sub New()
      MyBase.New(PropertyType.ecmDate)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmDate, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpDefaultValue As Nullable(Of DateTime))

      MyBase.New(PropertyType.ecmDate, lpName, lpSystemName)

      Try
        DefaultValue = lpDefaultValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace

