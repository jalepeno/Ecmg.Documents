' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  PropertyComparer.vb
'  Description :  [type_description_here]
'  Created     :  9/22/2011 9:20:32 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class PropertyComparer
    Inherits Comparer(Of IProperty)

#Region "Public Methods"

    Public Shared Function CompareProperties(x As IProperty, y As IProperty) As Integer
      Try

        Dim lintComparisonValue As Integer

        ' First check the name
        lintComparisonValue = String.Compare(x.Name, y.Name)

        If lintComparisonValue <> 0 Then
          ' The names are different, return the comparison value
          Return lintComparisonValue
        Else
          ' The names are the same, compare the system names
          lintComparisonValue = String.Compare(x.SystemName, y.SystemName)
        End If

        If lintComparisonValue <> 0 Then
          ' The names are different, return the comparison value
          Return lintComparisonValue
        Else
          ' The system names are the same, compare the type
          lintComparisonValue = x.Type.CompareTo(y.Type)
        End If

        If lintComparisonValue <> 0 Then
          ' The types are different, return the comparison value
          Return lintComparisonValue
        Else
          ' The types are the same, compare the cardinality.
          lintComparisonValue = x.Cardinality.CompareTo(y.Cardinality)
        End If

        If lintComparisonValue <> 0 Then
          ' The cardinalities are different, return the comparison value
          Return lintComparisonValue
        Else
          ' The cardinalities are the same, compare the values
          lintComparisonValue = x.HasValue.CompareTo(y.HasValue)
        End If

        If lintComparisonValue <> 0 Then
          ' One property has a value and the other does not, return the comparison value
          Return lintComparisonValue
        Else
          Select Case x.Cardinality
            Case Cardinality.ecmSingleValued
              lintComparisonValue = String.Compare(x.Value.ToString, y.Value.ToString)
            Case Cardinality.ecmMultiValued
              ' TODO: Implement multivalued property set comparisons.
          End Select
        End If

        Return lintComparisonValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Overriden Methods"

    Public Overrides Function Compare(x As IProperty, y As IProperty) As Integer
      Try

        Return CompareProperties(x, y)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace