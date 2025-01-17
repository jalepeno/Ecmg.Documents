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

  Public Class ClassificationStringProperty
    Inherits ClassificationProperty

#Region "Class Variables"

    Private mintmaxLength As Nullable(Of Integer)

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The maximum length (in characters) allowed for a value of this property.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If an application attempts to set the value of this property to a string 
    ''' larger than the specified maximum length, the repository MUST throw a constraint exception.  
    ''' </remarks>
    Public Property MaxLength As Nullable(Of Integer)
      Get
        Return mintmaxLength
      End Get
      Set(ByVal value As Nullable(Of Integer))
        Try
          If Type <> PropertyType.ecmString AndAlso value IsNot Nothing Then
            Throw New ConstraintException(String.Format(
              "MaxLength not valid for {0} property type, use this property only for String property types.",
              Type.ToString))
          End If
          mintmaxLength = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads Property DefaultValue As String
      Get
        Return MyBase.DefaultValue
      End Get
      Set(ByVal value As String)
        MyBase.DefaultValue = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Friend Sub New()
      MyBase.New(PropertyType.ecmString)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmString, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpDefaultValue As String)

      MyBase.New(PropertyType.ecmString, lpName, lpSystemName)

      Try

        If Not String.IsNullOrEmpty(lpDefaultValue) Then
          DefaultValue = lpDefaultValue
        Else
          DefaultValue = String.Empty
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Friend Sub New(ByVal lpName As String,
                   ByVal lpSystemName As String,
                   ByVal lpDefaultValue As String,
                   ByVal lpMaxLength As Nullable(Of Integer))
      MyBase.New(PropertyType.ecmString, lpName, lpSystemName)
      Try

        If lpDefaultValue IsNot Nothing Then
          DefaultValue = lpDefaultValue
        Else
          DefaultValue = String.Empty
        End If

        MaxLength = lpMaxLength

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace