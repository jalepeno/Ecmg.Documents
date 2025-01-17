' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ItemException.vb
'  Description :  [type_description_here]
'  Created     :  2/28/2012 4:17:17 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Exceptions

  Public Class ItemException
    Inherits CtsException

#Region "Class Variables"

    Private mstrIdentifier As String = String.Empty
    Private mobjItem As Object = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property Identifier As String
      Get
        Return mstrIdentifier
      End Get
    End Property

    Public ReadOnly Property Item As Object
      Get
        Return mobjItem
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new ItemException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal identifier As String)
      Me.New(identifier, FormatMessage(identifier))
    End Sub

    ''' <summary>
    ''' Creates a new ItemException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <param name="innerException">The inner Exception to attach.</param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal identifier As String, ByVal innerException As Exception)
      Me.New(identifier, FormatMessage(identifier), innerException)
    End Sub

    ''' <summary>
    ''' Creates a new ItemException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <param name="message">The exception message.</param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal identifier As String, ByVal message As String)
      MyBase.New(message)
      Try
        mstrIdentifier = identifier
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates a new ItemException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <param name="message">The exception message.</param>
    ''' <param name="innerException">The inner Exception to attach.</param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal identifier As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
      Try
        mstrIdentifier = identifier
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates a new ItemException 
    ''' with the specified item reference.
    ''' </summary>
    ''' <param name="item">The item reference.</param>
    ''' <param name="message">The exception message.</param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal item As Object, ByVal message As String)
      MyBase.New(message)
      Try
        mobjItem = item
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates a new ItemException 
    ''' with the specified item reference.
    ''' </summary>
    ''' <param name="item">The item reference.</param>
    ''' <param name="message">The exception message.</param>
    ''' <param name="innerException">The inner Exception to attach.</param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal item As Object, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
      Try
        mobjItem = item
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Shared Function FormatMessage(ByVal identifier As String) As String
      Try
        Return String.Format("Item '{0}'", identifier)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace