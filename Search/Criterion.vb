'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Search

  ''' <summary>Fully describes a single criterion for use in a search.</summary>
  Public Class Criterion
    Inherits Data.Criterion

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String,
               Optional ByVal lpPropertyName As String = Nothing,
               Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty,
               Optional ByVal lpOperator As pmoOperator = pmoOperator.opEquals,
               Optional ByVal lpSetEvaluation As SetEvaluation = Core.SetEvaluation.seAnd) ', ByVal lpDataType As pmoDataType)

      MyBase.New(lpName, lpPropertyName, lpPropertyScope, lpOperator, lpSetEvaluation)

    End Sub

  End Class

End Namespace