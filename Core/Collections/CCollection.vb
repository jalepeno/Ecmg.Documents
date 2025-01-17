'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports Documents.Utilities

#End Region

Namespace Core

  ''' <summary>
  ''' Base collection class which all Cts Collections classes inherit from
  ''' </summary>
  ''' <typeparam name="T"></typeparam>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Partial Public MustInherit Class CCollection(Of T)
    Inherits ObservableCollection(Of T)
    Implements IDisposable

#Region "Public Events"

    Public Event CollectionSorted(ByVal sender As Object, ByVal e As System.Collections.Specialized.NotifyCollectionChangedEventArgs)
    Public Event ItemPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
    ''' Overridden so that we may manually call registered handlers and differentiate between those that do and don't require Action.Reset args.
    Public Shadows Event CollectionChanged As NotifyCollectionChangedEventHandler

#End Region

#Region "Class Variables"

    Protected mobjObjects As New List(Of T)
    'Flag used to prevent OnCollectionChanged from firing during a bulk operation like Add(IEnumerable<T>) and Clear()
    Private _SuppressCollectionChanged As Boolean = False
    Private mobjNotificationEvent As PropertyChangedEventArgs
    Dim mobjPropertyChangedHandler As New PropertyChangedEventHandler(AddressOf OnItemPropertyChanged)

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Specifies whether or not the collection is read-only
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property IsReadOnly() As Boolean
      Get
        Return False
      End Get
    End Property

    ''' <summary>Gets the collection of items.</summary>
    Public Shadows ReadOnly Property Items As IList(Of T)
      Get
        Return MyBase.Items
      End Get
    End Property

#Region "Item Overloads"

    ''' <summary>
    ''' Gets or sets the element at the specified index.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the element to get or set.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Overridable Shadows Property Item(ByVal index As Integer) As T
      Get
        Try
          ' ApplicationLogging.WriteLogEntry(String.Format("Enter {0}::CCollection(Of T)::Get_Item({1} As Integer)", Me.GetType.Name, index), TraceEventType.Verbose)
#If CTSDOTNET = 1 Then
            'SiAuto.Main.EnterMethod(Level.Debug, Me, "get_Item(ByVal index As Integer)")
#End If
          If (MyBase.Items.Count > 0) Then
            If index < MyBase.Items.Count Then
              Return MyBase.Item(index)
            Else
              Throw New IndexOutOfRangeException(String.Format("The requested index '{0}' is not available.  The maximum available index in the collection is {1}.",
                                                               index, MyBase.Items.Count - 1))
            End If
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("{0}::CCollection(Of T)::Get_Item({1} As Integer)", Me.GetType.Name, index))
        Finally
          'ApplicationLogging.WriteLogEntry(String.Format("Exit {0}::CCollection(Of T)::Get_Item({1} As Integer)", Me.GetType.Name, index), TraceEventType.Verbose)
#If CTSDOTNET = 1 Then
          'SiAuto.Main.LeaveMethod(Level.Debug, Me, "get_Item(ByVal index As Integer)")

#End If
        End Try
      End Get
      Set(ByVal value As T)
        Try
          'ApplicationLogging.WriteLogEntry(String.Format("Enter {0}::CCollection(Of T)::Set_Item({1} As Integer)", Me.GetType.Name, index), TraceEventType.Verbose)
#If CTSDOTNET = 1 Then
         ' SiAuto.Main.EnterMethod(Level.Debug, Me, "set_Item(ByVal index As Integer)")

#End If
          If (MyBase.Items.Count > 0) Then
            If index < MyBase.Items.Count Then
              MyBase.Item(index) = value
            Else
              Throw New IndexOutOfRangeException(String.Format("The requested index '{0}' is not available.  The maximum available index in the collection is {1}.",
                                                               index, MyBase.Items.Count - 1))
            End If
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("{0}::CCollection(Of T)::Set_Item({1} As Integer)", Me.GetType.Name, index))
        Finally
          'ApplicationLogging.WriteLogEntry(String.Format("Exit {0}::CCollection(Of T)::Set_Item({1} As Integer)", Me.GetType.Name, index), TraceEventType.Verbose)
#If CTSDOTNET = 1 Then
          'SiAuto.Main.LeaveMethod(Level.Debug, Me, "set_Item(ByVal index As Integer)")

#End If
        End Try
      End Set
    End Property

    Public Overridable Function GetItemByIndex(ByVal index As Integer) As T
      Try
        Return Item(index)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Sub SetItemByIndex(ByVal index As Integer, ByVal value As T)
      Try
        Item(index) = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Function GetItemByName(ByVal name As String) As T
      Try
        Return Item(name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Sub SetItemByName(ByVal name As String, ByVal value As T)
      Try
        Item(name) = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>Gets the element using the specified name.</summary>
    ''' <param name="name">The name to look for</param>
    ''' <remarks>Assumes the item has a property called name or the type of the property is string,
    ''' else if it does not then the method will fail.</remarks>
    ''' <value></value>
    ''' <returns></returns>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Overridable Shadows Property Item(ByVal name As String) As T
      Get
        Try
          If Items.Count = 0 Then
            'ApplicationLogging.WriteLogEntry( _
            '  String.Format("Item '{0}' does not exist in the {1} collection, the collection is empty", _
            '                name, Me.GetType.Name), _
            '  Reflection.MethodBase.GetCurrentMethod, TraceEventType.Verbose, 61403)
            Return Nothing
          End If
          For Each lobjItem As Object In Items 'mobjObjects

            ' Added to deal with strings in the ccollection
            If (lobjItem.GetType.Name = "String") Then
              If String.Equals(lobjItem, name, StringComparison.InvariantCultureIgnoreCase) Then
                Return lobjItem
              End If
            Else
              ' Check to see if the items contained here support the 'Name' property
              If Helper.ObjectContainsProperty(lobjItem, "Name") Then
                If String.Equals(lobjItem.Name, name, StringComparison.InvariantCultureIgnoreCase) Then
                  Return lobjItem
                End If
              End If
              If Helper.ObjectContainsProperty(lobjItem, "Id") Then
                If String.Equals(lobjItem.Id, name, StringComparison.InvariantCultureIgnoreCase) Then
                  Return lobjItem
                End If
              End If
            End If

          Next

          ' We were unable to find the item
          '' Make an entry in the log and return nothing
          ''ApplicationLogging.WriteLogEntry( _
          ''  String.Format("Item '{0}' does not exist in the {1} collection", name, Me.GetType.Name), _
          ''  Reflection.MethodBase.GetCurrentMethod, TraceEventType.Verbose, 61404)
          Return Nothing
        Catch InvalidOpEx As InvalidOperationException
          ApplicationLogging.LogException(InvalidOpEx, Reflection.MethodBase.GetCurrentMethod)
          Return Nothing
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Get
      Set(ByVal value As T)
        Try
          Dim lobjItem As Object
          ' Added to deal with strings in the ccollection
          If (value.GetType.Name = "String") Then
            For lintCounter As Integer = 0 To MyBase.Count - 1
              lobjItem = CType(Item(lintCounter), T)
              If String.Equals(lobjItem, name, StringComparison.InvariantCultureIgnoreCase) Then
                MyBase.Item(lintCounter) = value
                Exit Property
              End If
            Next
            Throw New Exception("There is no Item by the Name '" & name & "'.")
          Else
            For lintCounter As Integer = 0 To MyBase.Count - 1
              lobjItem = CType(Item(lintCounter), T)
              If String.Equals(lobjItem.Name, name, StringComparison.InvariantCultureIgnoreCase) Then
                MyBase.Item(lintCounter) = value
                Exit Property
              End If
            Next
            Throw New Exception("There is no Item by the Name '" & name & "'.")
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>Gets the element using the specified name, ignoring or honoring its case.</summary>
    ''' <param name="name">The name to look for</param>
    ''' <param name="ignoreCase">A Syste.Boolean indicating a case-sensitive or insensitive comparison. 
    ''' (True indicates a case-insensitive comparison.)</param>
    ''' <remarks>Assumes the item has a property called name or the type of the property is string, 
    ''' elseif it does not then the method will fail.</remarks>
    ''' <value></value>
    ''' <returns></returns>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Overridable Shadows ReadOnly Property Item(ByVal name As String, ByVal ignoreCase As Boolean) As T
      Get
        For Each lobjObject As Object In Items 'mobjObjects

          'Added to deal with strings in the ccollection
          If (lobjObject.GetType.Name = "String") Then

            If ignoreCase Then
              If String.Equals(lobjObject, name, StringComparison.InvariantCultureIgnoreCase) Then
                Return lobjObject
              End If
            Else
              If String.Equals(lobjObject, name) Then
                Return lobjObject
              End If
            End If
          Else
            If ignoreCase Then
              If String.Equals(lobjObject.Name, name, StringComparison.InvariantCultureIgnoreCase) Then
                Return lobjObject
              End If
            Else
              If String.Equals(lobjObject.Name, name) Then
                Return lobjObject
              End If
            End If
          End If
        Next
      End Get
    End Property

#End Region

#End Region

#Region "Public Methods"

    ''' <summary>Adds a range of objects to the collection.</summary>
    Public Overridable Overloads Sub AddRange(ByVal items As IEnumerable(Of T))
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
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>Determines if an object exists in the collection with the specified name.</summary>
    ''' <param name="name">The object name to test.</param>
    ''' <returns>True if an object exists by the specified name, otherwise false.</returns>
    Public Overridable Overloads Function Contains(ByVal name As String) As Boolean
      Try

        Dim lobjItem As Object = Item(name)

        If lobjItem IsNot Nothing Then
          Return True
        Else
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Overloads Function Contains(ByVal item As T) As Boolean
      Try
        Return MyBase.Contains(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>Sorts the collection using the supplied comparer.</summary>
    Public Overridable Sub Sort(ByVal comparer As System.Collections.IComparer)
      Try
        Dim lobjOuterItem As Object
        Dim lobjInnerItem As Object
        Dim lintCurrentItemCounter As Integer
        Dim lintInnerSortCounter As Integer
        Dim lintMinMaxIndex As Integer
        Dim lobjMinMaxValue As Object
        'Dim lobjValue As Object
        Dim lblnSortCondition As Boolean
        Dim lobjComparableObject As IComparable
        Dim lintComparableReturn As Integer

        ' Make sure there is more than one 
        ' item before attempting to sort
        If Me.Count < 2 Then
          Exit Sub
        End If

        ' Make sure the items implement IComparable
        lobjComparableObject = TryCast(Item(0), IComparable)
        If lobjComparableObject Is Nothing Then
          Dim lstrErrorMessage As String = String.Format("Items of type {0} do not support sorting.  The class {0} does not implement IComparable.",
                                                         Item(0).GetType.Name)
          Throw New NotSupportedException(lstrErrorMessage)
        End If

        For lintCurrentItemCounter = 0 To Me.Count - 1
          lobjOuterItem = Item(lintCurrentItemCounter)

          'lobjMinMaxValue = CallByName(lobjItem, lpSortPropertyName, CallType.Get)
          lobjMinMaxValue = lobjOuterItem

          lintMinMaxIndex = lintCurrentItemCounter

          For lintInnerSortCounter = lintCurrentItemCounter + 1 To Me.Count - 1
            lobjInnerItem = Item(lintInnerSortCounter)
            'lobjValue = CallByName(lobjInnerItem, _
            ' lpSortPropertyName, CallType.Get)

            lobjComparableObject = TryCast(lobjInnerItem, IComparable)

            lintComparableReturn = comparer.Compare(lobjComparableObject, lobjOuterItem)
            'lintComparableReturn = lobjComparableObject.CompareTo(lobjOuterItem)

            If lintComparableReturn < 0 Then
              lblnSortCondition = True
            Else
              lblnSortCondition = False
            End If

            If (lblnSortCondition) Then
              lobjOuterItem = lobjInnerItem
              lintMinMaxIndex = lintInnerSortCounter
            End If

            lobjInnerItem = Nothing

          Next lintInnerSortCounter

          If (lintMinMaxIndex <> lintCurrentItemCounter) Then
            lobjOuterItem = Item(lintMinMaxIndex)

            Me.RemoveAt(lintMinMaxIndex)
            Me.InsertItem(lintCurrentItemCounter, lobjOuterItem)

            lobjOuterItem = Nothing
          End If

          lobjOuterItem = Nothing
        Next lintCurrentItemCounter

        RaiseEvent CollectionSorted(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        Return String.Format("{0}: {1} Items", Me.GetType.Name, Count)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        'Throw
        Return "Comparison Property Collection"
      End Try
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
          For Each lobjItem As Object In Me.Items
            If lobjItem.GetType.IsAssignableFrom(GetType(IDisposable)) Then
              CType(lobjItem, IDisposable).Dispose()
            End If
          Next
        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region "Event Handlers"

    Protected Sub OnItemPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
      Try
        RaiseEvent ItemPropertyChanged(sender, e)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Overrides Sub OnCollectionChanged(e As NotifyCollectionChangedEventArgs)
      Try
        If Not _SuppressCollectionChanged Then
          MyBase.OnCollectionChanged(e)
          If e.NewItems IsNot Nothing Then
            Dim lobjNotifyItem As INotifyPropertyChanged
            For Each lobjItem As Object In e.NewItems
              If TypeOf lobjItem Is INotifyPropertyChanged Then
                lobjNotifyItem = lobjItem
                AddHandler lobjNotifyItem.PropertyChanged, AddressOf OnItemPropertyChanged
              End If
            Next
          End If
          RaiseEvent CollectionChanged(Me, e)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''CollectionViews raise an error when they are passed a NotifyCollectionChangedEventArgs that indicates more than
    ''one element has been added or removed. They prefer to receive a "Action=Reset" notification, but this is not suitable
    ''for applications in code, so we actually check the type we're notifying on and pass a customized event args.
    'Protected Overridable Sub OnCollectionChangedMultiItem(e As NotifyCollectionChangedEventArgs)
    '  Try
    '    Dim handlers As NotifyCollectionChangedEventHandler = Nothing 'Me.CollectionChanged
    '    If handlers IsNot Nothing Then
    '      For Each handler As NotifyCollectionChangedEventHandler In handlers.GetInvocationList()
    '        handler(Me, If(Not (TypeOf handler.Target Is ICollectionView), e, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset)))
    '      Next
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "Extended Collection Methods"

    Protected Overrides Sub ClearItems()
      Try
        If Me.Count = 0 Then
          Return
        End If

        Dim removed As New List(Of T)(Me)
        _SuppressCollectionChanged = True
        MyBase.ClearItems()
        _SuppressCollectionChanged = False
#If CTSDOTNET = 1 Then
        OnCollectionChangedMultiItem(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove, removed))
#End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Public Sub Add(toAdd As IEnumerable(Of T))
    '  If Me = toAdd Then
    '    Throw New Exception("Invalid operation. This would result in iterating over a collection as it is being modified.")
    '  End If

    '  _SuppressCollectionChanged = True
    '  For Each item As T In toAdd
    '    Add(item)
    '  Next
    '  _SuppressCollectionChanged = False
    '  OnCollectionChangedMultiItem(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, New List(Of T)(toAdd)))
    'End Sub

#End Region

  End Class

End Namespace