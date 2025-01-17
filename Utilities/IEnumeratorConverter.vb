'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Utilities

  ''' <summary>
  ''' Converts an enumerator of one class to another, usually used for interfaces.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class IEnumeratorConverter(Of T)
    Inherits List(Of T)
    Implements IEnumerable(Of T)
    Implements IEnumerator

#Region "Class Variables"

    'Private ReadOnly mintCurrentIndex As Integer = 0
    Private ReadOnly mobjSourceType As Type
    'Private ReadOnly mobjDestinationType As Type

#End Region

#Region "Constructors"

    Public Sub New(lpSourceCollection As IList, lpSourceType As Type, lpDestinationType As Type)
      Try
        If lpDestinationType.IsAssignableFrom(lpSourceType) = False Then
          Throw New InvalidCastException(String.Format("Unable to cast '{0}' to '{1}'",
                                                     lpSourceType.Name, lpDestinationType.Name), 54821)
        End If

        mobjSourceType = lpSourceType
        'mobjDestinationType = lpDestinationType
        If lpSourceCollection IsNot Nothing Then
          AddRange(lpSourceCollection)
        End If


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overridable Overloads Sub AddRange(items As IEnumerable(Of T))
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

    Public Overloads Sub Add(item As T)
      Try
        If Me.Contains(item) = False Then
          ' Try to add the item to the collection
          MyBase.Add(item)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of T)
      Try
        Return MyBase.AsEnumerable.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IEnumerable(Of T) Implementation"

    Public ReadOnly Property Current As Object Implements System.Collections.IEnumerator.Current
      Get
        Try
          Return DirectCast(MyBase.ToList, IEnumerator(Of T)).Current
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
      Try
        Return DirectCast(MyBase.ToList, IEnumerator(Of T)).MoveNext
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Reset() Implements System.Collections.IEnumerator.Reset
      Try
        DirectCast(MyBase.ToList, IEnumerator(Of T)).Reset()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace