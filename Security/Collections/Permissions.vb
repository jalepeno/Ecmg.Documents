' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Permissions.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 3:51:31 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Security

  Public Class Permissions
    Inherits CCollection(Of ItemPermission)
    Implements IPermissions

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IPermission)

#End Region

#Region "IPermissions Implementation"

    Public Overloads Function Contains(name As String) As Boolean Implements IPermissions.Contains
      Try
        Return MyBase.Contains(name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Property Item(name As String) As IPermission Implements IPermissions.Item
      Get
        Try
          Return MyBase.Item(name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IPermission)
        Try
          MyBase.Item(name) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads Sub Add(item As IPermission) Implements System.Collections.Generic.ICollection(Of IPermission).Add
      Try
        MyBase.Add(CType(item, ItemPermission))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub AddRange(permissions As IPermissions) Implements IPermissions.AddRange
      Try
        MyBase.AddRange(permissions)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements System.Collections.Generic.ICollection(Of IPermission).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IPermission) As Boolean Implements System.Collections.Generic.ICollection(Of IPermission).Contains
      Try
        Return Contains(CType(item, ItemPermission))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IPermission, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IPermission).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IPermission).Count
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

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IPermission).IsReadOnly
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

    Public Overloads Function Remove(item As IPermission) As Boolean Implements System.Collections.Generic.ICollection(Of IPermission).Remove
      Try
        Return Remove(CType(item, ItemPermission))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IPermission) Implements System.Collections.Generic.IEnumerable(Of IPermission).GetEnumerator
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

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IPermission)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IPermission)(Me.ToArray, GetType(ItemPermission), GetType(IPermission))
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