'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  FileDoesNotExistException.vb
'   Description :  [type_description_here]
'   Created     :  5/12/2014 2:11:47 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Exceptions

  Public Class FileDoesNotExistException
    Inherits ItemDoesNotExistException

#Region "Constructors"

    Public Sub New(ByVal fileId As String)
      MyBase.New(fileId)
    End Sub

    Public Sub New(ByVal fileId As String, ByVal innerException As Exception)
      MyBase.New(fileId, innerException)
    End Sub

    Public Sub New(ByVal fileId As String, ByVal message As String)
      MyBase.New(fileId, message)
    End Sub

    Public Sub New(ByVal fileId As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(fileId, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shadows Function FormatMessage(ByVal fileId As String) As String
      Try
        Return String.Format("File '{0}' could not be found", fileId)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace