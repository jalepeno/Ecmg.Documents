' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FolderAlreadyExistsException.vb
'  Description :  [type_description_here]
'  Created     :  2/28/2012 4:25:08 PM
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

  Public Class FolderAlreadyExistsException
    Inherits ItemAlreadyExistsException

#Region "Constructors"

    Public Sub New(ByVal folderId As String)
      MyBase.New(folderId)
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

    Public Shared Function Create(parentFolderId As String, folderName As String, existingFolderId As String) As FolderAlreadyExistsException
      Try
        Return New FolderAlreadyExistsException(existingFolderId,
          String.Format("A folder already exists with parent folder id '{0}' and folder name '{1}'.", parentFolderId, folderName))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#Region "Private Methods"

    Private Shadows Function FormatMessage(ByVal folderId As String) As String
      Try
        Return String.Format("Folder '{0}' already exists.", folderId)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace