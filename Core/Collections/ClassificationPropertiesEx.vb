'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports Documents.Core.ClassificationProperty
Imports Documents.Utilities

#End Region

Namespace Core

  Partial Public Class ClassificationProperties
    Implements ICloneable
    Implements ITableable

#Region "Public Methods"

    Public Function GetAllChoiceLists() As ChoiceLists.ChoiceLists
      Try

        Dim lobjChoiceLists As New ChoiceLists.ChoiceLists

        For Each lobjProperty As ClassificationProperty In Me
          If ((lobjProperty.ChoiceList IsNot Nothing) AndAlso
              (lobjChoiceLists.Contains(lobjProperty.ChoiceList.Name) = False)) Then

            lobjChoiceLists.Add(lobjProperty.ChoiceList)

          End If
        Next

        lobjChoiceLists.Sort()

        Return lobjChoiceLists

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ICloneable Implementation"

    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjProperties As New ClassificationProperties

      Try
        For Each lobjProperty As ClassificationProperty In Me
          lobjProperties.Add(lobjProperty.Clone)
        Next
        Return lobjProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "ITableable Implementation"

    Public Function ToDataTable() As System.Data.DataTable Implements ITableable.ToDataTable
      Try

        ' Create New DataTable
        Dim lobjDataTable As New DataTable(String.Format("tbl{0}", Me.GetType.Name))

        ' Create columns        
        With lobjDataTable.Columns
          .Add("Name", System.Type.GetType("System.String"))
          .Add("SystemName", System.Type.GetType("System.String"))
          .Add("Cardinality", System.Type.GetType("System.String"))
          .Add("Type", System.Type.GetType("System.String"))
          .Add("Required", System.Type.GetType("System.Boolean"))
          .Add("Hidden", System.Type.GetType("System.Boolean"))
          .Add("System", System.Type.GetType("System.Boolean"))
          .Add("Settability", System.Type.GetType("System.String"))
          .Add("ChoiceList", System.Type.GetType("System.String"))
          .Add("DefaultValue", System.Type.GetType("System.String"))
          .Add("MaxLength", System.Type.GetType("System.String"))
          .Add("MinValue", System.Type.GetType("System.String"))
          .Add("MaxValue", System.Type.GetType("System.String"))
        End With

        Dim lstrCardinality As String = String.Empty
        Dim lstrSettability As String = String.Empty
        Dim lstrType As String = String.Empty
        Dim lstrDefaultValue As String = String.Empty
        Dim lintMaxLength As Nullable(Of Integer)
        Dim lstrMaxLength As String = String.Empty
        Dim lstrMinValue As String = String.Empty
        Dim lstrMaxValue As String = String.Empty

        'SortCollection("Name", True)
        Sort()

        For Each lobjProperty As ClassificationProperty In Me

          ' Re-initialize variables
          lintMaxLength = Nothing
          lstrMaxLength = String.Empty
          lstrMinValue = String.Empty
          lstrMaxValue = String.Empty

          If lobjProperty.Cardinality = Core.Cardinality.ecmMultiValued Then
            lstrCardinality = "Multi Valued"
          Else
            lstrCardinality = "Single Valued"
          End If

          Select Case lobjProperty.Settability
            Case SettabilityEnum.READ_ONLY
              lstrSettability = "Read Only"
            Case SettabilityEnum.READ_WRITE
              lstrSettability = "Read Write"
            Case SettabilityEnum.SETTABLE_ONLY_BEFORE_CHECKIN
              lstrSettability = "Only Before Checkin"
            Case SettabilityEnum.SETTABLE_ONLY_ON_CREATE
              lstrSettability = "Only on Create"
          End Select

          lstrType = lobjProperty.Type.ToString.Replace("ecm", String.Empty)

          If lobjProperty.DefaultValue IsNot Nothing Then
            lstrDefaultValue = lobjProperty.DefaultValue.ToString
          End If

          If TypeOf lobjProperty Is ClassificationStringProperty Then
            lintMaxLength = CType(lobjProperty, ClassificationStringProperty).MaxLength
            If lintMaxLength = 0 Then
              lintMaxLength = Nothing
            End If
          Else
            lintMaxLength = Nothing
          End If

          If lintMaxLength.HasValue Then
            lstrMaxLength = lintMaxLength.Value.ToString
          Else
            lstrMaxValue = String.Empty
          End If

          If (TypeOf lobjProperty Is ClassificationLongProperty) OrElse
            (TypeOf lobjProperty Is ClassificationDoubleProperty) OrElse
            (TypeOf lobjProperty Is ClassificationDateTimeProperty) Then
            lstrMinValue = CType(lobjProperty, Object).MinValue
            lstrMaxValue = CType(lobjProperty, Object).MaxValue
          End If

          If String.IsNullOrEmpty(lstrMaxLength) Then
            lstrMaxLength = String.Empty
          End If

          If String.IsNullOrEmpty(lstrMinValue) Then
            lstrMinValue = String.Empty
          End If

          If String.IsNullOrEmpty(lstrMaxValue) Then
            lstrMaxValue = String.Empty
          End If

          lobjDataTable.Rows.Add(lobjProperty.Name, lobjProperty.SystemName, lstrCardinality,
                                     lstrType, lobjProperty.IsRequired, lobjProperty.IsHidden,
                                     lobjProperty.IsSystemProperty, lstrSettability,
                                     lobjProperty.ChoiceListName, lstrDefaultValue, lstrMaxLength,
                                     lstrMinValue, lstrMaxValue)

        Next

        Return lobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace