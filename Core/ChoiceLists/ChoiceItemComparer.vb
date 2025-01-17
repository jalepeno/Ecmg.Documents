'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core.ChoiceLists

  Public Class ChoiceItemComparer
    Implements System.Collections.Generic.IComparer(Of ChoiceItem)
    Implements System.Collections.IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

      Try
        Return Compare(CType(x, ChoiceItem), CType(y, ChoiceItem))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function Compare(ByVal x As ChoiceItem, ByVal y As ChoiceItem) As Integer Implements System.Collections.Generic.IComparer(Of ChoiceItem).Compare
      Try
        ' Two null objects are equal
        If (x Is Nothing) And (y Is Nothing) Then Return 0

        ' Any non-null object is greater than a null object
        If (x Is Nothing) Then Return 1
        If (y Is Nothing) Then Return -1

        If x.DisplayName < y.DisplayName Then
          Return -1
        Else
          Return 1
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function
  End Class

End Namespace