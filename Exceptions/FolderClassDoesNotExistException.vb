' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FolderClassDoesNotExistException.vb
'  Description :  [type_description_here]
'  Created     :  3/5/2012 1:14:44 PM
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

  Public Class FolderClassDoesNotExistException
    Inherits ItemDoesNotExistException

#Region "Constructors"

    Public Sub New(ByVal folderClassName As String)
      MyBase.New(folderClassName)
    End Sub

    Public Sub New(ByVal folderClassName As String, ByVal innerException As Exception)
      MyBase.New(folderClassName, innerException)
    End Sub

    Public Sub New(ByVal folderClassName As String, ByVal message As String)
      MyBase.New(folderClassName, message)
    End Sub

    Public Sub New(ByVal folderClassName As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(folderClassName, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shadows Function FormatMessage(ByVal folderClassName As String) As String
      Try
        Return String.Format("No folder class with the name '{0}' could be found.", folderClassName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace