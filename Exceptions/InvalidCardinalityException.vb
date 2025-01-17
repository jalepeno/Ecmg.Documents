'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core


#End Region

Namespace Exceptions

  Public Class InvalidCardinalityException
    Inherits PropertyException

#Region "Class Variables"

    Private menuExpectedCardinality As Cardinality
    Private mobjProperty As ECMProperty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ExpectedCardinality() As Cardinality
      Get
        Return menuExpectedCardinality
      End Get
    End Property

    Public ReadOnly Property ActualCardinality() As Cardinality
      Get
        If menuExpectedCardinality = Cardinality.ecmSingleValued Then
          Return Cardinality.ecmMultiValued
        Else
          Return Cardinality.ecmSingleValued
        End If
      End Get
    End Property

    Public ReadOnly Property InvalidProperty() As ECMProperty
      Get
        Return mobjProperty
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="lpExpectedCardinality">The cardinality expected</param>
    ''' <param name="lpInvalidProperty">The invalid property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpExpectedCardinality As Cardinality, ByVal lpInvalidProperty As ECMProperty)
      MyBase.New(String.Format("The property '{0}' is not a {1} property.",
                               lpInvalidProperty.Name, lpExpectedCardinality.ToString),
                             lpInvalidProperty)
      menuExpectedCardinality = lpExpectedCardinality
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyTypeException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="lpExpectedCardinality">The cardinality expected</param>
    ''' <param name="lpInvalidProperty">The invalid property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal lpExpectedCardinality As Cardinality, ByVal lpInvalidProperty As ECMProperty)
      MyBase.New(message, lpInvalidProperty)
      menuExpectedCardinality = lpExpectedCardinality
    End Sub

#End Region

  End Class

End Namespace