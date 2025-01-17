'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------



Namespace Data
  ''' <summary>Collection of Criterion objects.</summary>
  Public Class Criteria
    'Inherits Search.Criteria

    Inherits Core.CCollection(Of Criterion)

#Region "Overloads"

    'Default Shadows Property Item(ByVal Name As String) As Criterion
    '  Get
    '    Try
    '      Dim lobjCriterion As Criterion
    '      For lintCounter As Integer = 0 To MyBase.Count - 1
    '        lobjCriterion = CType(MyBase.Item(lintCounter), Criterion)
    '        If lobjCriterion.Name = Name Then
    '          Return lobjCriterion
    '        End If
    '      Next
    '      Throw New Exception("There is no Item by the Name '" & Name & "'.")
    '      'Throw New InvalidArgumentException
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      '  Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    '  Set(ByVal value As Criterion)
    '    Try
    '      Dim lobjCriterion As Criterion
    '      For lintCounter As Integer = 0 To MyBase.Count - 1
    '        lobjCriterion = CType(MyBase.Item(lintCounter), Criterion)
    '        If lobjCriterion.Name = Name Then
    '          MyBase.Item(lintCounter) = value
    '        End If
    '      Next
    '      Throw New Exception("There is no Item by the Name '" & Name & "'.")
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      '  Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

    'Default Shadows Property Item(ByVal Index As Integer) As Criterion
    '  Get
    '    Return MyBase.Item(Index)
    '  End Get
    '  Set(ByVal value As Criterion)
    '    MyBase.Item(Index) = value
    '  End Set
    'End Property

#Region "Finalizer"

    Protected Overrides Sub Finalize()
      MyBase.Finalize()
    End Sub

#End Region

#End Region

  End Class
End Namespace