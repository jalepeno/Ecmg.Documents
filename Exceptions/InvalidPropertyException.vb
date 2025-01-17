' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  InvalidPropertyException.vb
'  Description :  [type_description_here]
'  Created     :  3/5/2012 1:30:03 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Exceptions

  Public Class InvalidPropertyException
    Inherits PropertyException

#Region "Constructors"

    ''' <summary>
    ''' Creates a new InvalidPropertyException
    ''' </summary>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal [property] As IProperty)
      MyBase.New([property])
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyException
    ''' </summary>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal [property] As IProperty)
      MyBase.New(message, [property])
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="innerException">The originating exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
    End Sub

    ''' <summary>
    ''' Creates a new InvalidPropertyException
    ''' </summary>
    ''' <param name="message">The error message</param>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <param name="innerException">The originating exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal [property] As IProperty, ByVal innerException As Exception)
      MyBase.New(message, [property], innerException)
    End Sub

#End Region

  End Class

End Namespace