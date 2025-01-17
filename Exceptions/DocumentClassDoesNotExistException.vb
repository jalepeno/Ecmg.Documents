' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DocumentClassDoesNotExistException.vb
'  Description :  [type_description_here]
'  Created     :  3/4/2012 10:30:45 AM
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

  Public Class DocumentClassDoesNotExistException
    Inherits ItemDoesNotExistException

#Region "Constructors"

    Public Sub New(ByVal documentClassName As String)
      MyBase.New(documentClassName)
    End Sub

    Public Sub New(ByVal documentClassName As String, ByVal innerException As Exception)
      MyBase.New(documentClassName, innerException)
    End Sub

    Public Sub New(ByVal documentClassName As String, ByVal message As String)
      MyBase.New(documentClassName, message)
    End Sub

    Public Sub New(ByVal documentClassName As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(documentClassName, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shadows Function FormatMessage(ByVal documentClassName As String) As String
      Try
        Return String.Format("No document class with the name '{0}' could be found.", documentClassName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace