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

  Public Class ClassificationDoubleProperty
    Inherits ClassificationProperty

#Region "Class Variables"

    Private mdblMinValue As Nullable(Of Double)
    Private mdblMaxValue As Nullable(Of Double)
    'Private mintPrecision As Nullable(Of Integer)

#End Region

#Region "Attributes specific to Double Object-Type Property Definitions"

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
    Public Property MinValue As Nullable(Of Double)
      Get
        If Type <> PropertyType.ecmDouble Then
          Throw New ConstraintException(String.Format(
            "MinValue not valid for {0} property type, use this property only for Double property types.",
            Type.ToString))
        End If
        Return mdblMinValue
      End Get
      Set(ByVal value As Nullable(Of Double))
        If Type <> PropertyType.ecmDouble Then
          Throw New ConstraintException(String.Format(
            "MinValue not valid for {0} property type, use this property only for Double property types.",
            Type.ToString))
        End If
        mdblMinValue = value
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
    Public Property MaxValue As Nullable(Of Double)
      Get
        If Type <> PropertyType.ecmDouble Then
          Throw New ConstraintException(String.Format(
            "MaxValue not valid for {0} property type, use this property only for Double property types.",
            Type.ToString))
        End If
        Return mdblMaxValue
      End Get
      Set(ByVal value As Nullable(Of Double))
        If Type <> PropertyType.ecmDouble Then
          Throw New ConstraintException(String.Format(
            "MaxValue not valid for {0} property type, use this property only for Double property types.",
            Type.ToString))
        End If
        mdblMaxValue = value
      End Set
    End Property

    ' ''' <summary>
    ' ''' This is the precision in bits supported for values of this property.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property Precision As enumDecimalPrecision
    '  Get
    '    If PropertyType <> enumPropertyType.Integer Then
    '      Throw New ConstraintException(String.Format( _
    '        "Precision not valid for {0} property type, use this property only for Decimal property types.", _
    '        PropertyType.ToString))
    '    End If
    '    Return _Precision
    '  End Get
    '  Set(ByVal value As enumDecimalPrecision)
    '    If PropertyType <> enumPropertyType.Integer Then
    '      Throw New ConstraintException(String.Format( _
    '        "Precision not valid for {0} property type, use this property only for Decimal property types.", _
    '        PropertyType.ToString))
    '    End If
    '    _Precision = value
    '  End Set
    'End Property

    Public Overloads Property DefaultValue As Nullable(Of Double)

#End Region

#Region "Constructors"

    Friend Sub New()
      MyBase.New(PropertyType.ecmDouble)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmDouble, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpDefaultValue As Nullable(Of Double))

      MyBase.New(PropertyType.ecmDouble, lpName, lpSystemName)

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
                   ByVal lpDefaultValue As Nullable(Of Double),
                   ByVal lpMinValue As Nullable(Of Double),
                   ByVal lpMaxValue As Nullable(Of Double))
      MyBase.New(PropertyType.ecmDouble, lpName, lpSystemName)
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

