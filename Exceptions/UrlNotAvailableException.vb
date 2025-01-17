'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  UrlNotAvailableException.vb
'   Description :  [type_description_here]
'   Created     :  5/13/2014 7:57:51 AM
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

  Public Class UrlNotAvailableException
    Inherits ItemDoesNotExistException

#Region "Constructors"

    Public Sub New(ByVal url As String)
      MyBase.New(url)
    End Sub

    Public Sub New(ByVal url As String, ByVal innerException As Exception)
      MyBase.New(url, innerException)
    End Sub

    Public Sub New(ByVal url As String, ByVal message As String)
      MyBase.New(url, message)
    End Sub

    Public Sub New(ByVal url As String, ByVal message As String, ByVal innerException As Exception)
      MyBase.New(url, message, innerException)
    End Sub

#End Region

#Region "Private Methods"

    Private Shadows Function FormatMessage(ByVal url As String) As String
      Try
        Return String.Format("Url '{0}' could not be reached", url)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace