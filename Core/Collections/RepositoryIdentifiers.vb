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

  Public Class RepositoryIdentifiers
    Inherits List(Of RepositoryIdentifier)

#Region "Public Methods"

    Public Function GetRepositoryId(ByVal lpRepositoryName As String) As String
      Try
        For Each lobjRepositoryIdentifier As RepositoryIdentifier In Me
          If lobjRepositoryIdentifier.Name.Equals(lpRepositoryName) OrElse lobjRepositoryIdentifier.Id.Equals(lpRepositoryName) Then
            Return lobjRepositoryIdentifier.Id
          End If
        Next

        ' We could not find it.
        Return String.Empty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace