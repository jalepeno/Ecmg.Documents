'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Collection of Content objects.</summary>
  <Serializable()>
  Partial Public Class Contents
    Implements ICloneable

    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjContents As New Contents()

      Try
        For Each lobjContent As Content In Me
          lobjContents.Add(lobjContent.Clone)
        Next
        Return lobjContents
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Private Sub Contents_CollectionChanged(ByVal sender As Object, ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles Me.CollectionChanged
      Try
        If e.Action = Specialized.NotifyCollectionChangedAction.Add Then
          For Each lobjContent As Content In e.NewItems
            If lobjContent.ShouldGetAvailableMetadata = True Then
              lobjContent.GetAvailableMetadata()
            End If
          Next
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace