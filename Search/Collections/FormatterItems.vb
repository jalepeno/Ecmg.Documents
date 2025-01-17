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
Imports Documents.Utilities

#End Region

Namespace Search

  Public Class FormatterItems
    Inherits CCollection(Of FormatterItem)
    Implements INamedItems

#Region "Class Variables"

    Private WithEvents mobjAvailableFormatters As New ClassificationProperties
    Private mobjAllFormatters As New ClassificationProperties

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal availableFormatters As ClassificationProperties)
      Try
        mobjAvailableFormatters = availableFormatters
        InitializeAllFormatters()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Contains the list of available items to format.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AvailableFormatters() As ClassificationProperties
      Get
        Return mobjAvailableFormatters
      End Get
      Set(ByVal value As ClassificationProperties)
        mobjAvailableFormatters = value
      End Set
    End Property

    Public ReadOnly Property AllFormatters() As ClassificationProperties
      Get
        Return mobjAllFormatters
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

    Default Overridable Shadows ReadOnly Property Item(ByVal lpFormatterItem As ClassificationProperty) As FormatterItem
      Get
        For Each lobjSelectedFormatterItem As FormatterItem In Items 'mobjObjects
          'Added to deal with strings in the ccollection
          If (lobjSelectedFormatterItem.FormatterItem.Equals(lpFormatterItem)) Then
            Return lobjSelectedFormatterItem
          End If
        Next
        Return Nothing
      End Get
    End Property

    Default Shadows Property Item(ByVal index As Integer) As FormatterItem
      Get
        Return MyBase.Items(index)
      End Get
      Set(ByVal value As FormatterItem)
        MyBase.Items(index) = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Factory method for creating a new FormatterItem object
    ''' </summary>
    ''' <param name="lpFormatterItem">The item to format with, should be a ClassificationProperty.</param>
    ''' <returns>A new FormatterItem object.</returns>
    ''' <remarks>Defaults sort to none and alignment to left.</remarks>
    ''' <exception cref="ArgumentNullException">
    ''' If a null value is provided for lpFormatterItem then an ArgumentNullException will be thrown.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' If lpFormatterItem is not of type ClasificationProperty then an ArgumentException will be thrown.
    ''' </exception>
    Public Function Create(ByVal lpFormatterItem As ClassificationProperty) As FormatterItem
      Try
        Return Create(lpFormatterItem, SortOption.None, AlignmentOption.Left)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Factory method for creating a new FormatterItem object
    ''' </summary>
    ''' <param name="lpFormatterItem">The item to format with, should be a ClassificationProperty.</param>
    ''' <param name="lpSort">The sort option for the new FormatterItem.</param>
    ''' <param name="lpAlignment">The alignment option for the new FormatterItem.</param>
    ''' <returns>A new FormatterItem object.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">
    ''' If a null value is provided for lpFormatterItem then an ArgumentNullException will be thrown.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' If lpFormatterItem is not of type ClasificationProperty then an ArgumentException will be thrown.
    ''' </exception>
    Public Function Create(ByVal lpFormatterItem As ClassificationProperty,
                           ByVal lpSort As SortOption,
                           ByVal lpAlignment As AlignmentOption) As FormatterItem
      Try

        ' Validate the item parameter

        ' Make sure we do not have a null object reference
        If lpFormatterItem Is Nothing Then
          Throw New ArgumentNullException("lpFormatterItem", "A valid ClassificationProperty is required.")
        End If

        ' Make sure the item is of the correct type
        If TypeOf (lpFormatterItem) Is ClassificationProperty Then
          Dim lobjReturnItem As New FormatterItem(lpFormatterItem, AvailableFormatters, AllFormatters)

          ' Assign the other parameter info
          With lobjReturnItem
            .Sort = lpSort
            .Alignment = lpAlignment
          End With

          Return lobjReturnItem

        Else
          Dim lstrErrorMessage As String = String.Format(
            "Type {0} not expected, an item of type ClassificationProperty is required.", lpFormatterItem.GetType.Name)
          Throw New ArgumentException(lstrErrorMessage, "lpFormatterItem")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub Add(ByVal lpFormatterItem As ClassificationProperty)
      'MyBase.Add(items)
      Try
        MyBase.Add(New FormatterItem(lpFormatterItem, AvailableFormatters, AllFormatters))
        AvailableFormatters.Remove(lpFormatterItem.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal lpFormatterItems As IList)
      Try
        For Each lobjItem As ClassificationProperty In lpFormatterItems
          MyBase.Add(New FormatterItem(lobjItem, AvailableFormatters, AllFormatters))
          AvailableFormatters.Remove(lobjItem.Name)
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal lpFormatterItems As Core.ClassificationProperties)
      'MyBase.Add(items)
      Try
        For Each lobjItem As ClassificationProperty In lpFormatterItems
          MyBase.Add(New FormatterItem(lobjItem, AvailableFormatters, AllFormatters))
          AvailableFormatters.Remove(lobjItem.Name)
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Function Remove(ByVal lpFormatterItem As ClassificationProperty) As Boolean
      'MyBase.Add(items)
      Try
        'MyBase.Remove(New SelectedFormatterItem(Of ClassificationProperty)(lpFormatterItem))
        Dim lobjLocatedItem As FormatterItem
        Dim lblnReturn As Boolean

        ' First try using the object reference
        lobjLocatedItem = Item(lpFormatterItem)
        If lobjLocatedItem IsNot Nothing Then
          lblnReturn = MyBase.Remove(lobjLocatedItem)
          If lblnReturn = True Then
            AvailableFormatters.Add(lpFormatterItem)
          End If
          Return lblnReturn
        Else
          If Helper.ObjectContainsProperty(lpFormatterItem, "Name") Then
            lobjLocatedItem = Item(CType(lpFormatterItem, Object).Name)
            If lobjLocatedItem IsNot Nothing Then
              lblnReturn = MyBase.Remove(lobjLocatedItem)
              If lblnReturn = True Then
                AvailableFormatters.Add(lpFormatterItem)
              End If
              Return lblnReturn
            End If
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function Remove(ByVal lpFormatterItems As Core.ClassificationProperties) As Boolean
      Try
        Dim lblnReturn As Boolean = False
        Dim lblnFailed As Boolean = False

        For Each lobjItem As ClassificationProperty In lpFormatterItems
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

    Private Sub InitializeAllFormatters()
      Try

        ' Copy all of the available formatters 
        ' into the list of all formatters

        If AvailableFormatters IsNot Nothing AndAlso AvailableFormatters.Count > 0 Then
          mobjAllFormatters.Clear()
          mobjAllFormatters.AddRange(AvailableFormatters)
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
            If Helper.ObjectTypeContainsProperty(GetType(ClassificationProperty), "Name") Then
              If String.Equals(lobjObject.FormatterItem.Name, name, StringComparison.InvariantCultureIgnoreCase) Then
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

    Public Shadows Sub Add(ByVal resultColumn As String)
      Try
        Dim lobjFormatterItem As FormatterItem '(resultColumn, AvailableFormatters, AllFormatters)
        If AvailableFormatters.Contains(resultColumn) Then
          lobjFormatterItem = New FormatterItem(resultColumn, AvailableFormatters, AllFormatters)
          Add(lobjFormatterItem)
        Else
          Throw New Exceptions.PropertyDoesNotExistException(
            String.Format("Unable to create formatter item for property '{0}', the property does not exist in the available list of formatters.",
            resultColumn), resultColumn)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal item As FormatterItem)
      Try
        If TypeOf (item) Is FormatterItem Then
          MyBase.Add(item)
          AvailableFormatters.Remove(item.FormatterItem.Name)
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

    Public Shadows Sub Add(ByVal item As Object) Implements INamedItems.Add
      Try
        If TypeOf (item) Is FormatterItem Then
          MyBase.Add(item)
        ElseIf TypeOf (item) Is ClassificationProperty Then
          MyBase.Add(New FormatterItem(DirectCast(item, ClassificationProperty), AvailableFormatters, AllFormatters))
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

    Public Shadows Function Remove(ByVal lpFormatterItemName As String) As Boolean _
      Implements INamedItems.Remove
      'MyBase.Add(items)
      Try
        'MyBase.Remove(New SelectedFormatterItem(Of ClassificationProperty)(lpFormatterItem))
        Dim lobjLocatedItem As FormatterItem

        ' First try using the object reference
        lobjLocatedItem = ItemByName(lpFormatterItemName)
        If lobjLocatedItem IsNot Nothing Then
          Return MyBase.Remove(lobjLocatedItem)
        Else
          If Helper.ObjectTypeContainsProperty(GetType(ClassificationProperty), "Name") Then
            lobjLocatedItem = ItemByName(lpFormatterItemName)
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

    Private Sub mobjAvailableFormatters_CollectionChanged(ByVal sender As Object,
                                                       ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles mobjAvailableFormatters.CollectionChanged
      Try
        If e.Action = Specialized.NotifyCollectionChangedAction.Add Then

          'AllFormatters.Add(e.NewItems)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace