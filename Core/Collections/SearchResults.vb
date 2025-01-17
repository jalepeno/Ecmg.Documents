'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Collection of SearchResult objects.</summary>
  Public Class SearchResults
    'Inherits Search.SearchResults
    Inherits CCollection(Of SearchResult)

#Region "Overloads"

    Default Shadows Property Item(ByVal ID As String) As SearchResult
      Get
        Try
          Dim lobjSearchResult As SearchResult
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjSearchResult = CType(MyBase.Item(lintCounter), SearchResult)
            If lobjSearchResult.ID = ID Then
              Return lobjSearchResult
            End If
          Next
          Throw New Exception("There is no Item by the ID '" & ID & "'.")
          'Throw New InvalidArgumentException
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As SearchResult)
        Try
          Dim lobjSearchResult As SearchResult
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjSearchResult = CType(MyBase.Item(lintCounter), SearchResult)
            If lobjSearchResult.ID = ID Then
              MyBase.Item(lintCounter) = value
            End If
          Next
          Throw New Exception("There is no Item by the ID '" & ID & "'.")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As SearchResult
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As SearchResult)
        MyBase.Item(Index) = value
      End Set
    End Property

#End Region

  End Class
End Namespace