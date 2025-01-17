'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  MapListActionItem.vb
'   Description :  [type_description_here]
'   Created     :  5/5/2014 9:08:07 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class MapListActionItem
    Inherits ActionItem

#Region "Class Variables"

    Private mobjMapList As New MapList
    Private WithEvents mobjMapStringList As ObservableCollection(Of String)
    Private mstrDelimiter As String = ":"

#End Region

#Region "Public Properties"

    Public Property OriginalMapList As MapList
      Get
        Try
          Return mobjMapList
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As MapList)
        Try
          mobjMapList = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property MapList As ObservableCollection(Of String)
      Get
        Try
          Return mobjMapStringList
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As ObservableCollection(Of String))
        Try
          mobjMapStringList = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Delimiter As String
      Get
        Try
          Return mstrDelimiter
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrDelimiter = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        'If Parameters.Count = 0 Then
        '  Parameters = GetDefaultParameters()
        'End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpActionItem As IActionItem)
      MyBase.New(lpActionItem)
    End Sub

    Public Sub New(lpActionItem As IActionItem, lpMapList As MapList, lpDelimiter As String)
      MyBase.New(lpActionItem)
      Try
        If lpMapList Is Nothing Then
          Throw New ArgumentNullException("lpMapList")
        End If

        mobjMapList = lpMapList
        mstrDelimiter = lpDelimiter
        mobjMapStringList = Me.mobjMapList.GetMapListItems(mstrDelimiter)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Sub mobjMapStringList_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles mobjMapStringList.CollectionChanged
      Try
        Dim lobjMapListValues As Values = CType(Me.Parameters.Item("MapList").Values, Values)

        Select Case e.Action
          Case NotifyCollectionChangedAction.Add
            For Each lobjItem As Object In e.NewItems
              lobjMapListValues.Add(lobjItem)
            Next

          Case NotifyCollectionChangedAction.Remove
            lobjMapListValues.Remove(e.OldItems)

          Case NotifyCollectionChangedAction.Replace
            For lintItemCounter As Integer = 0 To e.NewItems.Count - 1
              lobjMapListValues.Remove(lobjMapListValues.GetItemByName(e.OldItems(lintItemCounter)))
              lobjMapListValues.Add(e.NewItems(lintItemCounter))
            Next

          Case NotifyCollectionChangedAction.Reset
            Beep()

        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace