'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Exceptions

  ''' <summary>
  ''' Exception thrown when an item can't be located with the specified identifier.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ItemDoesNotExistException
    Inherits ItemException

#Region "Constructors"

    ''' <summary>
    ''' Creates a new ItemDoesNotExistException 
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

#End Region

#Region "Protected Methods"

    Protected Shared Function FormatMessage(ByVal identifier As String) As String
      Try
        Return String.Format("Item '{0}' could not be found.", identifier)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace