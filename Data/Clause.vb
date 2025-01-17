'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Imports Documents.Core
Imports Documents.Utilities

Namespace Data

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Clause

#Region "Class Variables"

    Private mobjCriteria As New Criteria
    Private menuSetEvaluation As SetEvaluation = SetEvaluation.seAnd

#End Region

#Region "Public Properties"

    Public Property Criteria() As Criteria
      Get
        Return mobjCriteria
      End Get
      Set(ByVal value As Criteria)
        mobjCriteria = value
      End Set
    End Property

    Public Property SetEvaluation() As SetEvaluation
      Get
        Return menuSetEvaluation
      End Get
      Set(ByVal value As SetEvaluation)
        menuSetEvaluation = value
      End Set
    End Property

#End Region

#Region "Constructors"

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
      Try
        Me.Criteria.Add(lpCriterion)
        menuSetEvaluation = lpConjunction
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjDebugStringBuilder As New Text.StringBuilder

        If Criteria Is Nothing OrElse Criteria.Count = 0 Then
          Return "No Criteria"
        End If
        For lintCriterionCounter As Integer = 0 To Criteria.Count - 1
          If lintCriterionCounter < (Criteria.Count - 1) Then
            lobjDebugStringBuilder.AppendFormat("{0} {1} ", Criteria(lintCriterionCounter).DebuggerIdentifier,
                                                Criteria(lintCriterionCounter).SetEvaluation.ToString.Substring(2).ToUpper)
          Else
            lobjDebugStringBuilder.Append(Criteria(lintCriterionCounter).DebuggerIdentifier)
          End If
        Next

        Return String.Format("({0})", lobjDebugStringBuilder.ToString)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace