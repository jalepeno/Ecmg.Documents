'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Search
  ''' <summary>
  ''' Contains a set of search results returned from a search
  ''' </summary>
  ''' <remarks>If the search operation resulted in an exception the HasException property will 
  ''' have a value of True and the Exception property will contain a 
  ''' reference to the exception thrown by the search.</remarks>
  Public Class SearchResultSet
    Inherits Core.SearchResultSet

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpException As Exception)
      MyBase.New(lpException)
    End Sub

    Public Sub New(ByVal lpResults As SearchResults)
      MyBase.New(lpResults)
    End Sub

    Public Sub New(ByVal lpResults As SearchResults, ByVal lpException As Exception)
      MyBase.New(lpResults, lpException)
    End Sub

  End Class

End Namespace