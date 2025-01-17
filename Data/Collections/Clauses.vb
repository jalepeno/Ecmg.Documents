'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Data
  Public Class Clauses
    Inherits Core.CCollection(Of Clause)

#Region "Public methods"

    Public Overloads Sub Add(ByVal lpClause As Clause)
      Try
        If lpClause.Criteria Is Nothing OrElse lpClause.Criteria.Count = 0 Then
          Exit Sub
        End If
        If Not Contains(lpClause) Then
          MyBase.Add(lpClause)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Add(ByVal lpCriterion As Criterion, ByVal lpConjunction As SetEvaluation) As Clause

      Try

        Dim lobjClause As New Clause(lpCriterion)
        MyBase.Add(lobjClause)

        Return lobjClause

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetFirstClause() As Clause
      Try
        If Count = 0 Then
          MyBase.Add(New Clause)
        End If
        Return Item(0)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace