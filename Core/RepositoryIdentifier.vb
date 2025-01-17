'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------



#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class RepositoryIdentifier
    Inherits RepositoryObject

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpName As String, ByVal lpDisplayName As String)
      Try
        Id = lpId
        Name = lpName
        DisplayName = lpDisplayName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace