'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core.ChoiceLists

  ''' <summary>
  ''' A collection of ChoiceValue objects
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ChoiceValues
    Inherits CCollection(Of ChoiceItem)

    Public Overrides Function Contains(name As String) As Boolean
      Try
        For Each lobjChoiceItem As ChoiceItem In Me
          If String.Compare(lobjChoiceItem.Name, name) = 0 Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub Sort()
      Try
        MyBase.Sort(New ChoiceItemComparer)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace