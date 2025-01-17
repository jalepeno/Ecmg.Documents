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
Imports Documents.Utilities


#End Region

Namespace Search

  Public Class FormatterItem
    Implements INotifyPropertyChanged

#Region "Class Variables"

    Private lobjFormatterItem As ClassificationProperty = Nothing
    Private menuSort As SortOption = SortOption.None
    Private mobjAvailableSortOptions As New List(Of String)
    Private menuAlignment As AlignmentOption = AlignmentOption.Left
    Private mobjAvailableAlignmentOptions As New List(Of String)

    Private WithEvents mobjAvailableFormatters As New ClassificationProperties
    Private mobjAvailableFormatterNames As New List(Of String)
    Private mobjAllFormatters As New ClassificationProperties
    Private mintPriority As Integer

#End Region

#Region "Public Properties"

    Public Property FormatterItem() As ClassificationProperty
      Get
        Return lobjFormatterItem
      End Get
      Set(ByVal value As ClassificationProperty)
        lobjFormatterItem = value
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("FormatterItem")
      End Set
    End Property

#Region "Alignment Properties"

    Public Property Alignment() As AlignmentOption
      Get
        Return menuAlignment
      End Get
      Set(ByVal value As AlignmentOption)
        menuAlignment = value
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("Alignment")
      End Set
    End Property

    Public Property AlignmentName() As String
      Get
        Return Alignment.ToString
      End Get
      Set(ByVal value As String)
        Select Case value.ToLower
          Case "left", "l", String.Empty
            Alignment = AlignmentOption.Left
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("AlignmentName")
          Case "center", "c"
            Alignment = AlignmentOption.Center
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("AlignmentName")
          Case "right", "r"
            Alignment = AlignmentOption.Right
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("AlignmentName")
          Case "stretch", "s", "justify", "j"
            Alignment = AlignmentOption.Stretch
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("AlignmentName")
        End Select
      End Set
    End Property

    Public ReadOnly Property AvailableAlignmentOptions() As List(Of String)
      Get
        If mobjAvailableAlignmentOptions.Count = 0 Then
          ' Initialize the list
          With mobjAvailableAlignmentOptions
            .Add("Left")
            .Add("Center")
            .Add("Right")
            .Add("Stretch")
          End With
        End If
        Return mobjAvailableAlignmentOptions
      End Get
    End Property

#End Region

#Region "Sort Properties"

    Public Property Sort() As SortOption
      Get
        Return menuSort
      End Get
      Set(ByVal value As SortOption)
        menuSort = value
        ' Call OnPropertyChanged whenever the property is updated
        OnPropertyChanged("Sort")
      End Set
    End Property

    Public Property SortName() As String
      Get
        Return Sort.ToString
      End Get
      Set(ByVal value As String)
        Select Case value.ToLower
          Case "none", String.Empty
            Sort = SortOption.None
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("SortName")
          Case "ascending", "asc"
            Sort = SortOption.Ascending
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("SortName")
          Case "descending", "desc"
            Sort = SortOption.Descending
            ' Call OnPropertyChanged whenever the property is updated
            OnPropertyChanged("SortName")
        End Select
      End Set
    End Property

    Public ReadOnly Property AvailableSortOptions() As List(Of String)
      Get
        If mobjAvailableSortOptions.Count = 0 Then
          ' Initialize the list
          With mobjAvailableSortOptions
            .Add("None")
            .Add("Ascending")
            .Add("Descending")
          End With
        End If
        Return mobjAvailableSortOptions
      End Get
    End Property

#End Region

    ''' <summary>
    ''' Gets or sets the priority for this item in the list of formatter items.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' Each FormatterItem should have a unique priority 
    ''' value and all values should be in sequence without gaps.
    ''' </remarks>
    Public Property Priority As Integer
      Get
        Return mintPriority
      End Get
      Set(ByVal value As Integer)
        mintPriority = value
      End Set
    End Property

    ''' <summary>
    ''' Contains the list of available properties to format.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property AvailableFormatters() As ClassificationProperties
      Get
        Return mobjAvailableFormatters
      End Get
    End Property

    ''' <summary>
    ''' Contains the list of names of available properties to format.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property AvailableFormatterNames() As List(Of String)
      Get
        If mobjAvailableFormatterNames.Count = 0 Then
          If AvailableFormatters.Count > 0 Then
            For Each lobjFormatter As ClassificationProperty In AvailableFormatters
              If Helper.ObjectContainsProperty(lobjFormatter, "Name") Then
                mobjAvailableFormatterNames.Add(CType(lobjFormatter, Object).Name)
              End If
            Next
          End If
        End If
        Return mobjAvailableFormatterNames
      End Get
    End Property

    Public Property FormatterName() As String
      Get
        If FormatterItem IsNot Nothing AndAlso
          Helper.ObjectContainsProperty(FormatterItem, "Name") Then
          Return DirectCast(FormatterItem, Object).Name
        Else
          Return String.Empty
        End If
      End Get
      Set(ByVal value As String)
        If AvailableFormatters.Contains(value) Then
          FormatterItem = AvailableFormatters(value)
          ' Do we need to manually remove the value 
          ' from the AvailableFilters collection or 
          ' will it be handled automatically by 
          ' the parent collection?
        End If
      End Set
    End Property

    Public ReadOnly Property AllFormatters() As ClassificationProperties
      Get
        Return mobjAllFormatters
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
    ''' <remarks>Use FormatterItems.Create() instead.</remarks>
    Public Sub New()
      If Helper.IsDeserializationBasedCall = False Then
        Throw New InvalidOperationException("Direct construction not allowed, use FormatterItems.Create() instead.")
      End If
    End Sub

    Protected Friend Sub New(ByVal lpFormatterItem As ClassificationProperty,
                   ByVal lpAvailableFormatters As ClassificationProperties,
                   ByVal lpAllFormatters As ClassificationProperties)
      FormatterItem = lpFormatterItem
      mobjAvailableFormatters = lpAvailableFormatters
      mobjAllFormatters = lpAllFormatters
    End Sub

    Protected Friend Sub New(ByVal lpFormatterItem As ClassificationProperty,
                   ByVal lpAlignment As AlignmentOption,
                   ByVal lpSort As SortOption,
                   ByVal lpAvailableFormatters As ClassificationProperties,
                   ByVal lpAllFormatters As ClassificationProperties)
      FormatterItem = lpFormatterItem
      Alignment = lpAlignment
      Sort = lpSort
      mobjAvailableFormatters = lpAvailableFormatters
      mobjAllFormatters = lpAllFormatters
    End Sub

    Protected Friend Sub New(ByVal lpColumnName As String,
                   ByVal lpAvailableFormatters As ClassificationProperties,
                   ByVal lpAllFormatters As ClassificationProperties)
      Try

        mobjAvailableFormatters = lpAvailableFormatters
        mobjAllFormatters = lpAllFormatters

        If mobjAvailableFormatters.Contains(lpColumnName) = False Then
          Throw New Exceptions.PropertyDoesNotExistException(
            String.Format("Unable to create formatter item for property '{0}', the property does not exist in the available list of formatters.",
                          lpColumnName), lpColumnName)
        End If

        With Me
          .FormatterName = lpColumnName
          .Alignment = AlignmentOption.Left
          .Sort = SortOption.None
          FormatterItem = .AvailableFormatters(lpColumnName)
        End With

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
      RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

#End Region

  End Class

End Namespace
