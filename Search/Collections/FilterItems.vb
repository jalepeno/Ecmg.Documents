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

Imports Documents.Core
Imports Documents.Data.Criterion
Imports Documents.Utilities

#End Region

Namespace Search

  Public Class FilterItems(Of T)
    Inherits CCollection(Of FilterItem(Of T))
    Implements INamedItems

#Region "Class Variables"

    Private WithEvents mobjAvailableFilters As CCollection(Of T)
    Private mobjAllFilters As New List(Of T)

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal availableFilters As CCollection(Of T))
      Try
        mobjAvailableFilters = availableFilters
        InitializeAllFilters()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Contains the list of available items to filter with.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AvailableFilters() As CCollection(Of T)
      Get
        Return mobjAvailableFilters
      End Get
      Set(ByVal value As CCollection(Of T))
        mobjAvailableFilters = value
      End Set
    End Property

    Public ReadOnly Property AllFilters() As List(Of T)
      Get
        Return mobjAllFilters
      End Get
    End Property

    Public Overloads ReadOnly Property Count() As Integer
      Get
        Try
          Return MyBase.Count
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Count", Me.GetType.Name))
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Default Overridable Shadows ReadOnly Property Item(ByVal lpFilterItem As T) As FilterItem(Of T)
      Get
        For Each lobjSelectedFilterItem As FilterItem(Of T) In Items 'mobjObjects
          'Added to deal with strings in the ccollection
          If (lobjSelectedFilterItem.FilterItem.Equals(lpFilterItem)) Then
            Return lobjSelectedFilterItem
          End If
        Next
        Return Nothing
      End Get
    End Property

    Default Shadows Property Item(ByVal index As Integer) As FilterItem(Of T)
      Get
        Return MyBase.Items(index)
      End Get
      Set(ByVal value As FilterItem(Of T))
        MyBase.Items(index) = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Factory method for creating a new FilterItem object
    ''' </summary>
    ''' <param name="lpItem">The item to filter with, should be either a ClassificationProperty or a DocumentClass object.</param>
    ''' <returns>A new FilterItem object.</returns>
    ''' <remarks>Defaults operator to equals, value to null and filter view to editable.</remarks>
    ''' <exception cref="ArgumentNullException">
    ''' If a null value is provided for lpItem then an ArgumentNullException will be thrown.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' If lpItem is not of type ClasificationProperty or DocumentClass then an ArgumentException will be thrown.
    ''' </exception>
    Public Function Create(ByVal lpItem As T) As FilterItem(Of T)
      Try
        Return Create(lpItem, pmoOperator.opEquals, Nothing, FilterView.Editable)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Factory method for creating a new FilterItem object
    ''' </summary>
    ''' <param name="lpItem">The item to filter with, should be either a ClassificationProperty or a DocumentClass object.</param>
    ''' <param name="lpOperator">The operator to assign to the new FilterItem.</param>
    ''' <param name="lpValue">A value to assign to the new FilterItem.  If lpValue is null or an empty string then no value will be assigned.</param>
    ''' <param name="lpView">The FilterView to assign to the new FilterItem.</param>
    ''' <returns>A new FilterItem object.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">
    ''' If a null value is provided for lpItem then an ArgumentNullException will be thrown.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' If lpItem is not of type ClasificationProperty or DocumentClass then an ArgumentException will be thrown.
    ''' </exception>
    Public Function Create(ByVal lpItem As T,
                           ByVal lpOperator As pmoOperator,
                           ByVal lpValue As Object,
                           ByVal lpView As FilterView) As FilterItem(Of T)
      Try

        ' Validate the item parameter

        ' Make sure we do not have a null object reference
        If lpItem Is Nothing Then
          Throw New ArgumentNullException("lpItem", "A valid ClassificationProperty or DocumentClass property is required.")
        End If

        ' Make sure the item is of the correct type
        If TypeOf (lpItem) Is ClassificationProperty OrElse TypeOf (lpItem) Is DocumentClass Then
          Dim lobjReturnItem As New FilterItem(Of T)(lpItem, AvailableFilters, AllFilters)

          ' Assign the other parameter info
          With lobjReturnItem
            .Operator = lpOperator
            .View = lpView
            If lpValue IsNot Nothing OrElse lpValue.ToString.Length > 0 Then
              .Value = lpValue.ToString
            Else
              .Value = String.Empty
            End If
          End With

          Return lobjReturnItem

        Else
          Dim lstrErrorMessage As String = String.Format(
            "Type {0} not expected, an item of type ClassificationProperty or DocumentClass is required.", lpItem.GetType.Name)
          Throw New ArgumentException(lstrErrorMessage, "lpItem")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Factory method for creating a new FilterItem object
    ''' </summary>
    ''' <param name="lpCriterion">The criterion associated with the new FilterItem</param>
    ''' <returns>A new FilterItem object.</returns>
    ''' <remarks>Defaults view to editable.</remarks>
    ''' <exception cref="ArgumentNullException">
    ''' If a null value is provided for lpItem then an ArgumentNullException will be thrown.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' If lpItem is not of type ClasificationProperty or DocumentClass then an ArgumentException will be thrown.
    ''' </exception>
    Public Function Create(ByVal lpCriterion As Data.Criterion) As FilterItem(Of T)
      Try
        Dim lobjClassificationItem As T = AvailableFilters(lpCriterion.PropertyName)
        If lobjClassificationItem Is Nothing Then
          Throw New InvalidOperationException(
            String.Format("Unable to create new FilterItem.  The criterion '{0}' does not correspond to an available filter.",
                          lpCriterion.PropertyName))
        End If
        Return Create(lobjClassificationItem, lpCriterion.Operator, lpCriterion.Value, FilterView.Editable)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub Add(ByVal lpFilterItem As T)
      'MyBase.Add(items)
      Try
        MyBase.Add(New FilterItem(Of T)(lpFilterItem, AvailableFilters, AllFilters))
        AvailableFilters.Remove(lpFilterItem)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal lpFilterItems As IList)
      Try
        For Each lobjItem As T In lpFilterItems
          MyBase.Add(New FilterItem(Of T)(lobjItem, AvailableFilters, AllFilters))
          AvailableFilters.Remove(lobjItem)
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal lpFilterItems As Core.CCollection(Of T))
      'MyBase.Add(items)
      Try
        For Each lobjItem As T In lpFilterItems
          MyBase.Add(New FilterItem(Of T)(lobjItem, AvailableFilters, AllFilters))
          AvailableFilters.Remove(lobjItem)
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Function Remove(ByVal lpFilterItem As T) As Boolean
      'MyBase.Add(items)
      Try
        'MyBase.Remove(New SelectedFilterItem(Of T)(lpFilterItem))
        Dim lobjLocatedItem As FilterItem(Of T)
        Dim lblnReturn As Boolean
        Dim lstrItemName As String = String.Empty

        ' First try using the object reference
        lobjLocatedItem = Item(lpFilterItem)
        If lobjLocatedItem IsNot Nothing Then
          lblnReturn = MyBase.Remove(lobjLocatedItem)
          If lblnReturn = True Then
            AvailableFilters.Add(lpFilterItem)
          End If
          Return lblnReturn
        Else
          If Helper.ObjectContainsProperty(lpFilterItem, "Name") Then
            lstrItemName = CType(lpFilterItem, Object).Name
            If Me.Contains(lstrItemName) Then
              lobjLocatedItem = Item(CType(lpFilterItem, Object).Name)
              If lobjLocatedItem IsNot Nothing Then
                lblnReturn = MyBase.Remove(lobjLocatedItem)
                If lblnReturn = True Then
                  AvailableFilters.Add(lpFilterItem)
                End If
                Return lblnReturn
              End If
            End If
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function Remove(ByVal lpFilterItems As Core.CCollection(Of T)) As Boolean
      Try
        Dim lblnReturn As Boolean = False
        Dim lblnFailed As Boolean = False

        For Each lobjItem As T In lpFilterItems
          lblnReturn = Remove(lobjItem)
          If lblnReturn = False Then
            lblnFailed = True
          End If
        Next

        If lblnFailed Then
          Return False
        Else
          Return True
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Sub InitializeAllFilters()
      Try

        ' Copy all of the available filters 
        ' into the list of all filters

        If AvailableFilters IsNot Nothing AndAlso AvailableFilters.Count > 0 Then
          mobjAllFilters.Clear()
          mobjAllFilters.AddRange(AvailableFilters)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "INamedItems Implementation"

    ''' <summary>
    ''' Get's the element using the specified name
    ''' </summary>
    ''' <param name="name">The name to look for</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Assumes the item has a property called name or the type of the property is string,
    ''' else if it does not then the method will fail.</remarks>
    Public Property ItemByName(ByVal name As String) As Object Implements Core.INamedItems.ItemByName
      Get
        For Each lobjObject As Object In Items 'mobjObjects

          'Added to deal with strings in the ccollection
          If (lobjObject.GetType.Name = "String") Then
            If String.Equals(lobjObject, name, StringComparison.InvariantCultureIgnoreCase) Then
              Return lobjObject
            End If
          Else
            ' Check to see if the items contained here support the 'Name' property
            If Helper.ObjectTypeContainsProperty(GetType(T), "Name") Then
              If String.Equals(lobjObject.FilterItem.Name, name, StringComparison.InvariantCultureIgnoreCase) Then
                Return lobjObject
              End If
            End If
          End If

        Next

        Return Nothing

      End Get
      Set(ByVal value As Object)
        MyBase.Item(name) = value
      End Set
    End Property

    Public Shadows Sub Add(ByVal lpCriterion As Data.Criterion)
      Try
        MyBase.Add(Create(lpCriterion))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal item As FilterItem(Of T))
      Try
        MyBase.Add(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal item As Object) Implements INamedItems.Add
      Try
        If TypeOf (item) Is FilterItem(Of T) Then
          MyBase.Add(item)
        ElseIf TypeOf (item) Is T Then
          MyBase.Add(New FilterItem(Of T)(item, AvailableFilters, AllFilters))
        Else
          Throw New ArgumentException(
            String.Format("Invalid object type.  Item type of {0} not expected.",
                         item.GetType.Name), "item")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Function Remove(ByVal lpFilterItemName As String) As Boolean _
      Implements INamedItems.Remove
      'MyBase.Add(items)
      Try
        'MyBase.Remove(New SelectedFilterItem(Of T)(lpFilterItem))
        Dim lobjLocatedItem As FilterItem(Of T)

        ' First try using the object reference
        lobjLocatedItem = ItemByName(lpFilterItemName)
        If lobjLocatedItem IsNot Nothing Then
          Return MyBase.Remove(lobjLocatedItem)
        Else
          If Helper.ObjectTypeContainsProperty(GetType(T), "Name") Then
            lobjLocatedItem = ItemByName(lpFilterItemName)
            If lobjLocatedItem IsNot Nothing Then
              Return MyBase.Remove(lobjLocatedItem)
            End If
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    Private Sub mobjAvailableFilters_CollectionChanged(ByVal sender As Object,
                                                       ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles mobjAvailableFilters.CollectionChanged
      Try
        If e.Action = Specialized.NotifyCollectionChangedAction.Add Then

          'AllFilters.Add(e.NewItems)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace