' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  AuditEvents.vb
'  Description :  [type_description_here]
'  Created     :  3/13/2012 1:06:45 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  ''' <summary>Contains a collection of audit events.</summary>
  Public Class AuditEvents
    Inherits CCollection(Of AuditEvent)
    Implements IAuditEvents

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IAuditEvent)

#End Region

#Region "IAuditEvents Implementation"

    Public Overloads Function Contains(id As String) As Boolean Implements IAuditEvents.Contains
      Try
        Return MyBase.Contains(id)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Property Item(id As String) As IAuditEvent Implements IAuditEvents.Item
      Get
        Try
          Return MyBase.Item(id)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IAuditEvent)
        Try
          MyBase.Item(id) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads Sub Add(item As IAuditEvent) Implements System.Collections.Generic.ICollection(Of IAuditEvent).Add
      Try
        Add(CType(item, AuditEvent))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements System.Collections.Generic.ICollection(Of IAuditEvent).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IAuditEvent) As Boolean Implements System.Collections.Generic.ICollection(Of IAuditEvent).Contains
      Try
        Return Contains(CType(item, AuditEvent))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IAuditEvent, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IAuditEvent).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IAuditEvent).Count
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

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IAuditEvent).IsReadOnly
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

    Public Overloads Function Remove(item As IAuditEvent) As Boolean Implements System.Collections.Generic.ICollection(Of IAuditEvent).Remove
      Try
        Return Remove(CType(item, AuditEvent))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IAuditEvent) Implements System.Collections.Generic.IEnumerable(Of IAuditEvent).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IAuditEvent)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IAuditEvent)(Me.ToArray, GetType(AuditEvent), GetType(IAuditEvent))
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

  End Class

End Namespace