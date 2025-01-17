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

  Public Class PropertyException
    Inherits CtsException

#Region "Class Variables"

    Private mobjProperty As IProperty = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property [Property]() As Core.ECMProperty
      Get
        Return mobjProperty
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new PropertyException
    ''' </summary>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal [property] As IProperty)
      MyBase.New(String.Format("The property '{0}' was not expected.", [property].Name))
      mobjProperty = [property]
    End Sub

    ''' <summary>
    ''' Creates a new PropertyException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    ''' <summary>
    ''' Creates a new PropertyException
    ''' </summary>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal [property] As IProperty)
      MyBase.New(message)
      mobjProperty = [property]
    End Sub

    ''' <summary>
    ''' Creates a new PropertyException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="innerException">The originating exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
    End Sub

    ''' <summary>
    ''' Creates a new PropertyException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <param name="innerException">The originating exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal [property] As IProperty, ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mobjProperty = [property]
    End Sub

#End Region

  End Class

End Namespace