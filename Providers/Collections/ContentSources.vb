'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core



Namespace Providers
  ''' <summary>Collection of ContentSource objects.</summary>
  Public Class ContentSources
    Inherits CCollection(Of ContentSource)

#Region "Overloaded Methods"

    Public Overloads Function Exists(ByVal lpContentSourceName As String) As Boolean
      For Each lobjContentSource As ContentSource In Me
        If lobjContentSource.Name = lpContentSourceName Then
          Return True
        End If
      Next

      Return False

    End Function

    Public Overloads Sub Add(ByVal lpContentSourceConnectionString As String)
      MyBase.Add(New ContentSource(lpContentSourceConnectionString))
    End Sub

    Public Overloads Function Remove(ByVal lpContentSourceName As String) As Boolean
      For Each lobjContentSource As ContentSource In Me
        If lobjContentSource.Name = lpContentSourceName Then
          Return Remove(lobjContentSource)
        End If
      Next

      Return False

    End Function

#End Region

  End Class

End Namespace