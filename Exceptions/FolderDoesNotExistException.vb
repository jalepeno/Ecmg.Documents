' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FolderDoesNotExistException.vb
'  Description :  Exception thrown when a folder can't be located 
'                 with the specified identifier.
'  Created     :  2/28/2012 3:21:32 PM
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
  ''' Exception thrown when a folder can't be located with the specified identifier.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FolderDoesNotExistException
    Inherits ItemDoesNotExistException

#Region "Constructors"

    Public Sub New(ByVal folderId As String)
      Me.New(folderId, FormatMessage(folderId))
    End Sub

    Public Sub New(ByVal folderId As String, ByVal innerException As Exception)
      MyBase.New(folderId, innerException)
    End Sub

    Public Sub New(ByVal folderId As String, ByVal message As String)
      MyBase.New(folderId, message)
    End Sub

    Public Sub New(ByVal folderId As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(folderId, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shared Shadows Function FormatMessage(ByVal folderId As String) As String
      Try
        Return String.Format("No folder with the identifier '{0}' could be found", folderId)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace