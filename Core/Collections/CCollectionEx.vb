'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Documents.Utilities

#End Region

Namespace Core

  '<ComVisible(True), ComDefaultInterface(GetType(Collection))> _
  '<Serializable()> _
  '<DebuggerDisplay("Count = 0")> _
  '<DebuggerDisplay("CTS Generic Collection")> _
  '<DebuggerTypeProxy(GetType(CCollectionDebugView))> _

  <Serializable()>
  Public MustInherit Class CCollection(Of T)
    Inherits ObservableCollection(Of T)
    Implements ICollection(Of T)
    Implements INotifyCollectionChanged

    '#Region "Public Events"

    '    Public Event CollectionChanged(ByVal sender As Object, ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Implements System.Collections.Specialized.INotifyCollectionChanged.CollectionChanged
    ' Public Event CollectionSorted(ByVal sender As Object, ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs)

    '#End Region

#Region "Class Variables"

    Private mobjObjectArray As ArrayList
    Private mobjContainedObjectType As System.Type = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property ContainedObjectType As System.Type
      Get
        If mobjContainedObjectType Is Nothing Then
          mobjContainedObjectType = GetContainedObjectType()
        End If
        Return mobjContainedObjectType
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpCollection As IEnumerable(Of T))
      MyBase.New(lpCollection)
    End Sub

#End Region

#Region "Public Methods"

    Public Function GetContainedObjectType() As System.Type
      Try
        Dim lobjGenericArguments As System.Type()
        Dim lobjBaseType As System.Type = Me.GetType.BaseType
        If lobjBaseType IsNot Nothing Then
          If lobjBaseType.Name.StartsWith("CCollection") = False Then
            Do Until lobjBaseType.Name.StartsWith("CCollection") = True
              lobjBaseType = lobjBaseType.BaseType
            Loop
          End If
          lobjGenericArguments = lobjBaseType.GetGenericArguments
          If lobjGenericArguments.Length > 0 Then
            Return lobjGenericArguments.FirstOrDefault
          Else
            Return GetType(Object)
          End If
        Else
          Return GetType(Object)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub MoveDown(item As T)
      Try
        Dim lintIndex As Integer = Items.IndexOf(item)
        If lintIndex < Items.Count - 1 Then
          Move(lintIndex, lintIndex + 1)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub MoveUp(item As T)
      Try
        Dim lintIndex As Integer = Items.IndexOf(item)
        If lintIndex > 0 Then
          Move(lintIndex, lintIndex - 1)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToArray() As Object()
      Try
        mobjObjectArray = New ArrayList()
        Dim mobjType As System.Type = Me.ContainedObjectType

        If Me.Count > 0 Then
          For Each mobjItem As Object In Me
            mobjObjectArray.Add(mobjItem)
          Next
        End If

        Return mobjObjectArray.ToArray(mobjType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry(String.Format("Error converting CCollection(of T) {0} '{1}'", Me.GetType.Name, ex.Message), TraceEventType.Information)
      End Try
      Return Nothing
    End Function

    Public Overrides Function ToString() As String
      Try
        Return String.Format("{0}: Count = {1}", Me.GetType.Name, Count)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ICollection Implementation"

    '''' <summary>
    '''' Adds an item of the necessary class to the collection
    '''' </summary>
    '''' <param name="item">The item to add</param>
    '''' <remarks></remarks>
    'Public Overridable Sub Add(ByVal item As T) 'Implements ICollection(Of T).Add
    '  Try
    '    mobjObjects.Add(item)
    '    'RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, item, mobjObjects.Count - 1))
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Add('{1}')", Me.GetType.Name, item.ToString))
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ''' <summary>
    ''' Adds a collection of items to the collection
    ''' </summary>
    ''' <param name="items"></param>
    ''' <remarks></remarks>
    Public Overridable Overloads Sub Add(ByVal items As CCollection(Of T))
      Try
        ' Add each of the items in the collection
        For Each lobjItem As T In items
          Try
            If Me.Contains(lobjItem) = False Then
              ' Try to add the item to the collection
              Add(lobjItem)
            End If
          Catch ex As Exception
            ' We were unable to add the item, log an error and continue
            ApplicationLogging.WriteLogEntry(
              String.Format("Unable to add the item to the collection: '{0}'", ex.Message),
              TraceEventType.Error)
            Continue For
          End Try
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Overloads Sub Add(items As IEnumerable(Of T))
      Try
        If Me.Equals(items) Then
          Throw New Exception("Invalid operation. This would result in iterating over a collection as it is being modified.")
        End If

        _SuppressCollectionChanged = True

        ' Add each of the items in the collection
        For Each lobjItem As T In items
          Try
            If Me.Contains(lobjItem) = False Then
              ' Try to add the item to the collection
              Add(lobjItem)
            End If
          Catch ex As Exception
            ' We were unable to add the item, log an error and continue
            ApplicationLogging.WriteLogEntry(
              String.Format("Unable to add the item to the collection: '{0}'", ex.Message),
              TraceEventType.Error)
            Continue For
          End Try
        Next
        _SuppressCollectionChanged = False
        'OnCollectionChangedMultiItem(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add,
        'New List(Of T)(items)))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Overloads Sub Remove(ByVal items As IList)
      Try
        If Me.Equals(items) Then
          Throw New Exception("Invalid operation. This would result in iterating over a collection as it is being modified.")
        End If

        _SuppressCollectionChanged = True

        Dim lobjItem As T
        For lintCounter As Integer = items.Count - 1 To 0 Step -1
          Try
            lobjItem = items(lintCounter)
            If Me.Contains(lobjItem) Then
              ' Try to remove the item from the collection
              Remove(lobjItem)
            End If
          Catch ex As Exception
            ' We were unable to remove the item, log an error and continue
            ApplicationLogging.WriteLogEntry(
              String.Format("Unable to remove the item from the collection: '{0}'", ex.Message),
              TraceEventType.Error)
            Continue For
          End Try
        Next

        _SuppressCollectionChanged = False
        Dim lobjNotificationList As New List(Of T)
        For Each Item As T In items
          lobjNotificationList.Add(Item)
        Next
        'OnCollectionChangedMultiItem(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove, _
        '                                                                  New List(Of T)(items)))

        'OnCollectionChangedMultiItem(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove,
        '                                                          lobjNotificationList))

        'For Each lobjItem As T In items
        '  Try
        '    If Me.Contains(lobjItem) = True Then
        '      ' Try to remove the item from the collection
        '      Remove(lobjItem)
        '    End If
        '  Catch ex As Exception
        '    ' We were unable to add the item, log an error and continue
        '    ApplicationLogging.WriteLogEntry( _
        '      String.Format("Unable to remove the item from the collection: '{0}'", ex.Message), _
        '      TraceEventType.Error)
        '    Continue For
        '  End Try
        'Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Overloads Sub Remove(ByVal items As CCollection(Of T))
      Try
        ' Remove each of the items from the collection
        For Each lobjItem As T In items
          Try
            If Me.Contains(lobjItem) = True Then
              ' Try to remove the item from the collection
              Remove(lobjItem)
            End If
          Catch ex As Exception
            ' We were unable to add the item, log an error and continue
            ApplicationLogging.WriteLogEntry(
              String.Format("Unable to remove the item from the collection: '{0}'", ex.Message),
              TraceEventType.Error)
            Continue For
          End Try
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    '''' <summary>
    '''' Clears all items from the collection.
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub Clear() Implements ICollection(Of T).Clear
    '  Try
    '    mobjObjects.Clear()
    '    RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Clear", Me.GetType.Name))
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    '''' <summary>
    '''' Determines whether or not a specific item is contained in the collection.
    '''' </summary>
    '''' <param name="item"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Contains(ByVal item As T) As Boolean Implements ICollection(Of T).Contains
    '  Try
    '    Return mobjObjects.Contains(item)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Contains('{1}')", Me.GetType.Name, item.ToString))
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Determines whether or not a specific item is contained in the collection based on the item name.
    '''' </summary>
    '''' <param name="name">The name to compare to</param>
    '''' <returns></returns>
    '''' <remarks>Assumes the item has a property called name, if it does not then the method will fail.</remarks>
    'Public Overridable Function Contains(ByVal name As String) As Boolean
    '  Try
    '    For Each lobjObject As Object In mobjObjects
    '      If String.Compare(lobjObject.Name, name, True) = 0 Then
    '        Return True
    '      End If
    '    Next

    '    Return False

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Copies the collection of items to the specified array.
    '''' </summary>
    '''' <param name="array"></param>
    '''' <param name="arrayIndex"></param>
    '''' <remarks></remarks>
    'Public Sub CopyTo(ByVal array() As T, ByVal arrayIndex As Integer) Implements ICollection(Of T).CopyTo
    '  Try
    '    mobjObjects.CopyTo(array, arrayIndex)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::CopyTo(array(), '{1}')", Me.GetType.Name, arrayIndex))
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ' ''' <summary>
    ' ''' The number of items present in the collection.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Shadows ReadOnly Property Count() As Long 'Implements ICollection(Of T).Count
    '  Get
    '    'Dim lintCount As Integer = 0
    '    Try
    '      'If Items Is Nothing Then
    '      '  Return lintCount
    '      'Else
    '      '  For lintCounter As Integer = 0 To Items.Count
    '      '    lintCount += 1
    '      '  Next
    '      '  Return lintCount 'Items.Count
    '      'End If
    '      'Return MyBase.Count ' mobjObjects.Count

    '      Return Items.Count

    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Count", Me.GetType.Name))
    '      'Return lintCount
    '    End Try
    '  End Get
    'End Property

    'Private ReadOnly Property DebuggerDisplay As String
    '  Get
    '    'Return String.Format("{0}: Count = {1}", Me.GetType.Name, 0) 'Count.ToString)
    '    Return "Count = 25"
    '    'Return 0
    '  End Get
    'End Property

    '''' <summary>
    '''' Removes the specified item from the collection.
    '''' </summary>
    '''' <param name="item"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Remove(ByVal item As T) As Boolean Implements ICollection(Of T).Remove
    '  Try
    '    mobjObjects.Remove(item)
    '    RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove, item))
    '    Return True
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Remove('{1}')", Me.GetType.Name, item.ToString))
    '    Return False
    '  End Try
    'End Function

    'Public Sub RemoveAt(ByVal index As Integer)
    '  Try
    '    If index < 0 Then
    '      Throw New ArgumentOutOfRangeException(String.Format("index value of {0} is less than zero.", index))
    '    End If

    '    mobjObjects.RemoveAt(index)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    '''' <summary>
    '''' Returns an enumerator that iterates through a CCollection(Of T) object
    '''' </summary>
    '''' <returns>A CCollection(Of T).Enumerator object for the CCollection(Of T) object</returns>
    '''' <remarks></remarks>
    'Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
    '  Try
    '    Return mobjObjects.GetEnumerator
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::GetEnumerator", Me.GetType.Name))
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
    '  Try
    '    Return mobjObjects.GetEnumerator
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::GetEnumerator1", Me.GetType.Name))
    '    Return Nothing
    '  End Try
    'End Function

#End Region

#Region "Overridable Properties and Methods"

    ''' <summary>
    ''' Checks to see if an item exists with the specified index number.
    ''' </summary>
    ''' <param name="index">The index number to check for.</param>
    ''' <returns>True if the index number is valid, otherwise False.</returns>
    ''' <remarks></remarks>
    Public Overridable Function ItemExists(ByVal index As Integer) As Boolean
      Try
        If (MyBase.Items.Count > 0) Then
          If index < MyBase.Items.Count Then
            Return True
          Else
            Return False
          End If
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Sort()
      Try
        Sort(Comparer.Default)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Public Overridable Sub Sort(ByVal comparer As System.Collections.IComparer)
    '  Try
    '    Dim lobjOuterItem As Object
    '    Dim lobjInnerItem As Object
    '    Dim lintCurrentItemCounter As Integer
    '    Dim lintInnerSortCounter As Integer
    '    Dim lintMinMaxIndex As Integer
    '    Dim lobjMinMaxValue As Object
    '    'Dim lobjValue As Object
    '    Dim lblnSortCondition As Boolean
    '    Dim lobjComparableObject As IComparable
    '    Dim lintComparableReturn As Integer

    '    ' Make sure there is more than one 
    '    ' item before attempting to sort
    '    If Me.Count = 1 Then
    '      Exit Sub
    '    End If

    '    ' Make sure the items implement IComparable
    '    lobjComparableObject = TryCast(Item(0), IComparable)
    '    If lobjComparableObject Is Nothing Then
    '      Dim lstrErrorMessage As String = String.Format("Items of type {0} do not support sorting.  The class {0} does not implement IComparable.", _
    '                                                     Item(0).GetType.Name)
    '      Throw New NotSupportedException(lstrErrorMessage)
    '    End If

    '    For lintCurrentItemCounter = 0 To Me.Count - 1
    '      lobjOuterItem = Item(lintCurrentItemCounter)

    '      'lobjMinMaxValue = CallByName(lobjItem, lpSortPropertyName, CallType.Get)
    '      lobjMinMaxValue = lobjOuterItem

    '      lintMinMaxIndex = lintCurrentItemCounter

    '      For lintInnerSortCounter = lintCurrentItemCounter + 1 To Me.Count - 1
    '        lobjInnerItem = Item(lintInnerSortCounter)
    '        'lobjValue = CallByName(lobjInnerItem, _
    '        ' lpSortPropertyName, CallType.Get)

    '        lobjComparableObject = TryCast(lobjInnerItem, IComparable)

    '        lintComparableReturn = comparer.Compare(lobjComparableObject, lobjOuterItem)
    '        'lintComparableReturn = lobjComparableObject.CompareTo(lobjOuterItem)

    '        If lintComparableReturn < 0 Then
    '          lblnSortCondition = True
    '        Else
    '          lblnSortCondition = False
    '        End If

    '        If (lblnSortCondition) Then
    '          lobjOuterItem = lobjInnerItem
    '          lintMinMaxIndex = lintInnerSortCounter
    '        End If

    '        lobjInnerItem = Nothing

    '      Next lintInnerSortCounter

    '      If (lintMinMaxIndex <> lintCurrentItemCounter) Then
    '        lobjOuterItem = Item(lintMinMaxIndex)

    '        Me.RemoveAt(lintMinMaxIndex)
    '        Me.InsertItem(lintCurrentItemCounter, lobjOuterItem)

    '        lobjOuterItem = Nothing
    '      End If

    '      lobjOuterItem = Nothing
    '    Next lintCurrentItemCounter
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Sub SortCollection(ByVal lpSortPropertyName As String, ByVal lpAscending As Boolean)
      Try
        Dim lobjItem As Object
        Dim lintCurrentItemCounter As Integer
        Dim j As Integer
        Dim lintMinMaxIndex As Integer
        Dim lobjMinMaxValue As Object
        Dim lobjValue As Object
        Dim lblnSortCondition As Boolean

        For lintCurrentItemCounter = 1 To Me.Count - 1
          lobjItem = Item(lintCurrentItemCounter)

          lobjMinMaxValue = CallByName(lobjItem, lpSortPropertyName, CallType.Get)

          lintMinMaxIndex = lintCurrentItemCounter

          For j = lintCurrentItemCounter + 1 To Me.Count - 1
            lobjItem = Item(j)
            lobjValue = CallByName(lobjItem,
             lpSortPropertyName, CallType.Get)

            If (lpAscending) Then
              lblnSortCondition = (lobjValue < lobjMinMaxValue)
            Else
              lblnSortCondition = (lobjValue > lobjMinMaxValue)
            End If

            If (lblnSortCondition) Then
              lobjMinMaxValue = lobjValue
              lintMinMaxIndex = j
            End If

            lobjItem = Nothing
          Next j

          If (lintMinMaxIndex <> lintCurrentItemCounter) Then
            lobjItem = Item(lintMinMaxIndex)

            Me.RemoveAt(lintMinMaxIndex)
            Me.InsertItem(lintCurrentItemCounter, lobjItem)

            lobjItem = Nothing
          End If

          lobjItem = Nothing
        Next lintCurrentItemCounter

        RaiseEvent CollectionSorted(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub


    '''' <summary>
    '''' Sorts the elements in the entire collection using the System.IComparable implementation of each element.
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub Sort()
    '  Try
    '    mobjObjects.Sort()
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Sort", Me.GetType.Name))
    '  End Try
    'End Sub

    '''' <summary>
    '''' Sorts the elements in the entire collection using the specified comparer.
    '''' </summary>
    '''' <param name="Comparer">The System.Collections.IComparer implementation to use when comparing elements, -or- null to use the System.IComparable implementation of each element.</param>
    '''' <remarks></remarks>
    'Public Sub Sort(ByVal Comparer As Collections.Generic.IComparer(Of T))
    '  Try
    '    mobjObjects.Sort(Comparer)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CCollection::Sort({1})", Me.GetType.Name, Comparer.GetType.Name))
    '  End Try
    'End Sub

#End Region

  End Class

End Namespace