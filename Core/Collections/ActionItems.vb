'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ActionItems.vb
'   Description :  [type_description_here]
'   Created     :  4/24/2014 2:14:18 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ActionItems
    Inherits CCollection(Of ActionItem)
    Implements IActionItems
    Implements IDisposable

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IActionItem)
    Private mblnDisposedValue As Boolean ' To detect redundant calls

#End Region

#Region "Public Events"

    Public Shadows Event ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements IActionItems.ItemPropertyChanged

#End Region

#Region "Public Properties"

    Public Shadows Property Item(name As String) As IActionItem Implements IActionItems.Item
      Get
        Try
          Return MyBase.Item(name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IActionItem)
        Try
          MyBase.Item(name) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads ReadOnly Property Count As Integer Implements Generic.ICollection(Of IActionItem).Count
      Get
        Try
          Return MyBase.Count
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements Generic.ICollection(Of IActionItem).IsReadOnly
      Get
        Try
          Return MyBase.IsReadOnly
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Private Properties"

    Protected ReadOnly Property IsDisposed() As Boolean
      Get
        Try
          Return mblnDisposedValue
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IActionItem)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IActionItem)(Me.ToArray, GetType(ActionItem), GetType(IActionItem))
          End If
          Return mobjEnumerator
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Methods"

    'Public Shared Function Create(lpItems As Object) As IActionItems
    '  Try
    '    If lpItems Is Nothing Then
    '      Throw New ArgumentNullException("lpItems")
    '    End If

    '    If Not TypeOf lpItems Is IEnumerable Then
    '      Throw New ArgumentException("lpItems is not enumerable.")
    '    End If

    '    Dim lobjItems As IActionItems = New ActionItems

    '    For Each lobjItem As Object In DirectCast(lpItems, IEnumerable)
    '      If TypeOf lobjItem Is IActionItem Then
    '        lobjItems.Add(lobjItem)
    '      End If
    '    Next

    '    Return lobjItems

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Overloads Sub AddRange(lpItems As IActionItems) Implements IActionItems.AddRange
      Try
        For Each lobjItem As IActionItem In lpItems
          If TypeOf lobjItem Is ActionItem Then
            MyBase.Add(CType(lobjItem, ActionItem))
          Else
            MyBase.Add(New ActionItem(lobjItem))
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function Contains(name As String) As Boolean Implements IActionItems.Contains
      Try
        Return MyBase.Contains(name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function GetItemByIndex(index As Integer) As IActionItem Implements IActionItems.GetItemByIndex
      Try
        Return MyBase.GetItemByIndex(index)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function GetItemByName(name As String) As IActionItem Implements IActionItems.GetItemByName
      Try
        Return MyBase.GetItemByName(name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub SetItemByIndex(index As Integer, value As IActionItem) Implements IActionItems.SetItemByIndex
      Try
        MyBase.SetItemByIndex(index, value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub SetItemByName(name As String, value As IActionItem) Implements IActionItems.SetItemByName
      Try
        MyBase.SetItemByName(name, value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(item As IActionItem) Implements Generic.ICollection(Of IActionItem).Add
      Try
        Add(CType(item, ActionItem))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(item As ActionItem)
      Try
        If Contains(item) = False Then
          MyBase.Add(item)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements Generic.ICollection(Of IActionItem).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IActionItem) As Boolean Implements Generic.ICollection(Of IActionItem).Contains
      Try
        Return Contains(CType(item, ActionItem))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Contains(item As ActionItem) As Boolean
      Try
        Return MyBase.Contains(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IActionItem, arrayIndex As Integer) Implements Generic.ICollection(Of IActionItem).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Remove(item As IActionItem) As Boolean Implements Generic.ICollection(Of IActionItem).Remove
      Try
        Return Remove(CType(item, ActionItem))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Remove(item As ActionItem) As Boolean
      Try
        Return MyBase.Remove(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As Generic.IEnumerator(Of IActionItem) Implements Generic.IEnumerable(Of IActionItem).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Clone() As Object Implements ICloneable.Clone

      Dim lobjItems As New ActionItems

      Try
        For Each lobjItem As ActionItem In Me
          lobjItems.Add(lobjItem.Clone)
        Next
        Return lobjItems
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "IDisposable Support"

    ' IDisposable
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      If Not Me.mblnDisposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
          MyBase.Dispose(disposing)
        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.mblnDisposedValue = True
    End Sub

#End Region

  End Class

End Namespace