'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Header"

' Copyright © 2009-2010 Conteage Corp
' All code contained herein is copyrighted by ECMG
' Enterprise Content Management Group, LLC
' 
' Code may not be used or reproduced in any form 
' without explicit written permission.

#End Region

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Core
Imports Documents.Data.Criterion
Imports Documents.Utilities


#End Region

Namespace Search

  Public Class FilterItem(Of T)
    Implements INotifyPropertyChanged

#Region "Class Variables"

    Private lobjFilterItem As T = Nothing
    Private menuView As FilterView = FilterView.Editable
    Private mobjAvailableViews As New List(Of String)
    Private menuOperator As pmoOperator = pmoOperator.opEquals
    Private mobjAvailableOperators As New List(Of String)
    Private mstrValue As String = String.Empty
    Private WithEvents mobjAvailableFilters As CCollection(Of T)
    Private mobjAvailableFilterNames As New List(Of String)
    Private mobjAllFilters As List(Of T)

#End Region

#Region "Public Properties"

    Public Property FilterItem() As T
      Get
        Return lobjFilterItem
      End Get
      Set(ByVal value As T)
        lobjFilterItem = value
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("FilterItem")
      End Set
    End Property

    Public Property View() As FilterView
      Get
        Return menuView
      End Get
      Set(ByVal value As FilterView)
        menuView = value
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("View")
      End Set
    End Property

    Public WriteOnly Property View(ByVal name As String) As FilterView
      Set(ByVal value As FilterView)
        If String.Equals(name, "ReadOnly", StringComparison.InvariantCultureIgnoreCase) Then
          View = FilterView.ReadOnly
        ElseIf String.Equals(name, "Editable", StringComparison.InvariantCultureIgnoreCase) Then
          View = FilterView.Editable
        ElseIf String.Equals(name, "Required", StringComparison.InvariantCultureIgnoreCase) Then
          View = FilterView.Required
        ElseIf String.Equals(name, "Hidden", StringComparison.InvariantCultureIgnoreCase) Then
          View = FilterView.Hidden
        End If
      End Set
    End Property

    Public Property ViewName() As String
      Get
        Return View.ToString
      End Get
      Set(ByVal value As String)
        Select Case value.ToLower
          Case "editable"
            View = FilterView.Editable

          Case "hidden"
            View = FilterView.Hidden

          Case "readonly"
            View = FilterView.ReadOnly

          Case "required"
            View = FilterView.Required

          Case Else
            View = FilterView.Editable

        End Select

        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("ViewName")

      End Set
    End Property

    Public ReadOnly Property AvailableViews() As List(Of String)
      Get
        If mobjAvailableViews.Count = 0 Then
          ' Initialize the list
          With mobjAvailableViews
            .Add("Editable")
            .Add("Hidden")
            .Add("ReadOnly")
            .Add("Required")
          End With
        End If
        Return mobjAvailableViews
      End Get
    End Property

    Public Property [Operator]() As pmoOperator
      Get
        Return menuOperator
      End Get
      Set(ByVal value As pmoOperator)
        menuOperator = value
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("Operator")
      End Set
    End Property

    Public Property OperatorName() As String
      Get
        Return [Operator].ToString.Remove(0, 2)
      End Get
      Set(ByVal value As String)
        Select Case value.ToLower
          Case "equals", "="
            [Operator] = pmoOperator.opEquals
          Case "notequal", "<>"
            [Operator] = pmoOperator.opNotEqual
          Case "greaterthan", ">"
            [Operator] = pmoOperator.opGreaterThan
          Case "lessthan", "<"
            [Operator] = pmoOperator.opLessThan
          Case "lessthanorequalto", "<="
            [Operator] = pmoOperator.opLessThanOrEqualTo
          Case "greaterthanorequalto", ">="
            [Operator] = pmoOperator.opGreaterThanOrEqualTo
          Case "like", "%", "*"
            [Operator] = pmoOperator.opLike
          Case "notlike"
            [Operator] = pmoOperator.opNotLike
          Case "in"
            [Operator] = pmoOperator.opIn
          Case "startswith"
            [Operator] = pmoOperator.opStartsWith
          Case "endswith"
            [Operator] = pmoOperator.opEndsWith
          Case "contentcontains"
            [Operator] = pmoOperator.opContentContains
          Case Else
            Exit Property
        End Select
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("OperatorName")
      End Set
    End Property

    Public ReadOnly Property AvailableOperators() As List(Of String)
      Get
        If mobjAvailableOperators.Count = 0 Then
          ' Initialize the list
          With mobjAvailableOperators
            .Add("Equals")
            .Add("LessThan")
            .Add("GreaterThan")
            .Add("LessThanOrEqualTo")
            .Add("GreaterThanOrEqualTo")
            .Add("Like")
            .Add("NotLike")
            .Add("In")
            .Add("EndsWith")
            .Add("ContentContains")
            .Add("StartsWith")
            .Add("NotEqual")
          End With
        End If
        Return mobjAvailableOperators
      End Get
    End Property

    ''' <summary>
    ''' The value to filter with, the criterion value.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Value() As String
      Get
        Return mstrValue
      End Get
      Set(ByVal value As String)
        mstrValue = value
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("Value")
      End Set
    End Property

    ''' <summary>
    ''' Contains the list of available items to filter with.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property AvailableFilters() As CCollection(Of T)
      Get
        Return mobjAvailableFilters
      End Get
    End Property

    ''' <summary>
    ''' Contains the list of names of available items to filter with.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property AvailableFilterNames() As List(Of String)
      Get
        Try
          If mobjAvailableFilterNames.Count = 0 Then
            If AvailableFilters.Count > 0 Then
              For Each lobjFilter As T In AvailableFilters
                If Helper.ObjectContainsProperty(lobjFilter, "Name") Then
                  mobjAvailableFilterNames.Add(CType(lobjFilter, Object).Name)
                End If
              Next
            End If
          End If
          Return mobjAvailableFilterNames
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property FilterName() As String
      Get
        Try
          If FilterItem IsNot Nothing AndAlso
                   Helper.ObjectContainsProperty(FilterItem, "Name") Then
            Return DirectCast(FilterItem, Object).Name
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          If AvailableFilters.Contains(value) Then
            FilterItem = AvailableFilters(value)
            ' Do we need to manually remove the value 
            ' from the AvailableFilters collection or 
            ' will it be handled automatically by 
            ' the parent collection?
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("FilterName")
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property AllFilters() As List(Of T)
      Get
        Return mobjAllFilters
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' For deserialization purposes only, if called directly will throw an InvalidOperationException.
    ''' </summary>
    ''' <exception cref="InvalidOperationException">
    ''' Unless called as a result of deserialization will throw InvalidOperationException.
    ''' </exception>
    ''' <remarks>Use FilterItems.Create() instead.</remarks>
    Public Sub New()
      Try
        If Helper.IsDeserializationBasedCall = False Then
          Throw New InvalidOperationException("Direct construction not allowed, use FilterItems.Create() instead.")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpFilterItem As T,
                   ByVal lpAvailableFilters As CCollection(Of T),
                   ByVal lpAllFilters As List(Of T))
      Try
        FilterItem = lpFilterItem
        mobjAvailableFilters = lpAvailableFilters
        mobjAllFilters = lpAllFilters
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Friend Sub New(ByVal lpFilterItem As T,
                   ByVal lpView As FilterView,
                   ByVal lpAvailableFilters As CCollection(Of T),
                   ByVal lpAllFilters As List(Of T))
      Try
        FilterItem = lpFilterItem
        View = lpView
        mobjAvailableFilters = lpAvailableFilters
        mobjAllFilters = lpAllFilters
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "INotifyPropertyChanged Implementation"

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    ' Create the OnPropertyChanged method to raise the event
    Protected Sub OnPropertyChanged(ByVal name As String)
      Try
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace