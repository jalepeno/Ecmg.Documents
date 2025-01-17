'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Search

  Public Class Clause
    Inherits Data.Clause

    'Private mobjCriteria As New Criteria
    'Private menuSetEvaluation As SetEvaluation = SetEvaluation.seAnd

    'Public Property Criteria() As Criteria
    '  Get
    '    Return mobjCriteria
    '  End Get
    '  Set(ByVal value As Criteria)
    '    mobjCriteria = value
    '  End Set
    'End Property

    'Public Property SetEvaluation() As SetEvaluation
    '  Get
    '    Return menuSetEvaluation
    '  End Get
    '  Set(ByVal value As SetEvaluation)
    '    menuSetEvaluation = value
    '  End Set
    'End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpCriterion As Criterion)
      Try
        Me.Criteria.Add(lpCriterion)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpCriterion As Criterion, ByVal lpConjunction As SetEvaluation)
      MyBase.New(lpCriterion, lpConjunction)
      'Try
      '  Me.Criteria.Add(lpCriterion)
      '  menuSetEvaluation = lpConjunction
      'Catch ex As Exception
      '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  ' Re-throw the exception to the caller
      '  Throw
      'End Try
    End Sub

  End Class
End Namespace