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
  ''' Exception thrown when a document can't be located with the specified identifier.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentDoesNotExistException
    Inherits ItemDoesNotExistException

#Region "Constructors"

    Public Sub New(ByVal docId As String)
      MyBase.New(docId)
    End Sub

    Public Sub New(ByVal docId As String, ByVal innerException As Exception)
      MyBase.New(docId, innerException)
    End Sub

    Public Sub New(ByVal docId As String, ByVal message As String)
      MyBase.New(docId, message)
    End Sub

    Public Sub New(ByVal docId As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(docId, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shadows Function FormatMessage(ByVal docId As String) As String
      Try
        Return String.Format("No document with the identifier '{0}' could be found", docId)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace