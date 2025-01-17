' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ItemAlreadyExistsException.vb
'  Description :  Exception thrown when a proposed new item already exists 
'                 with the specified identifier.
'  Created     :  2/28/2012 3:44:13 PM
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

  ''' <summary>
  ''' Exception thrown when a proposed new item already exists with the specified identifier.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ItemAlreadyExistsException
    Inherits ItemException

#Region "Constructors"

    ''' <summary>
    ''' Creates a new ItemAlreadyExistsException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal identifier As String)
      Me.New(identifier, FormatMessage(identifier))
    End Sub

    ''' <summary>
    ''' Creates a new ItemDoesNotExistException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <param name="innerException">The inner Exception to attach.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal identifier As String, ByVal innerException As Exception)
      Me.New(identifier, FormatMessage(identifier), innerException)
    End Sub

    ''' <summary>
    ''' Creates a new ItemDoesNotExistException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <param name="message">The exception message.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal identifier As String, ByVal message As String)
      MyBase.New(identifier, message)
    End Sub

    ''' <summary>
    ''' Creates a new ItemDoesNotExistException 
    ''' with the specified identifier.
    ''' </summary>
    ''' <param name="identifier">The identifier for the item.</param>
    ''' <param name="message">The exception message.</param>
    ''' <param name="innerException">The inner Exception to attach.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal identifier As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(identifier, message, innerException)
    End Sub

    ''' <summary>
    ''' Creates a new ItemDoesNotExistException 
    ''' with the specified item reference.
    ''' </summary>
    ''' <param name="item">The item reference.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal item As Object)
      MyBase.New(item, "Item already exists.")
    End Sub

    ''' <summary>
    ''' Creates a new ItemDoesNotExistException 
    ''' with the specified item reference.
    ''' </summary>
    ''' <param name="item">The item reference.</param>
    ''' <param name="message">The exception message.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal item As Object, ByVal message As String)
      MyBase.New(item, message)
    End Sub

    ''' <summary>
    ''' Creates a new ItemDoesNotExistException 
    ''' with the specified item reference.
    ''' </summary>
    ''' <param name="item">The item reference.</param>
    ''' <param name="message">The exception message.</param>
    ''' <param name="innerException">The inner Exception to attach.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal item As Object, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(item, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shared Function FormatMessage(ByVal identifier As String) As String
      Try
        Return String.Format("Item '{0}' already exists.", identifier)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace