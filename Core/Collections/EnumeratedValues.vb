' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  EnumeratedValues.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 10:28:26 AM
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

  Public Class EnumeratedValues
    Inherits CCollection(Of EnumeratedValue)
    Implements IEnumeratedValues

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IEnumeratedValue)

#End Region

    Public Overloads Function Contains(name As String) As Boolean Implements IEnumeratedValues.Contains
      Try
        Return MyBase.Contains(name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Property Item(name As String) As IEnumeratedValue Implements IEnumeratedValues.Item
      Get
        Try
          Return MyBase.Item(name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IEnumeratedValue)
        Try
          MyBase.Item(name) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads Sub Add(item As IEnumeratedValue) Implements System.Collections.Generic.ICollection(Of IEnumeratedValue).Add
      Try
        Add(CType(item, EnumeratedValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Clear() Implements System.Collections.Generic.ICollection(Of IEnumeratedValue).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IEnumeratedValue) As Boolean Implements System.Collections.Generic.ICollection(Of IEnumeratedValue).Contains
      Try
        Return Contains(CType(item, EnumeratedValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IEnumeratedValue, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IEnumeratedValue).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IEnumeratedValue).Count
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

    Public Shadows ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IEnumeratedValue).IsReadOnly
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

    Public Overloads Function Remove(item As IEnumeratedValue) As Boolean Implements System.Collections.Generic.ICollection(Of IEnumeratedValue).Remove
      Try
        Return Remove(CType(item, EnumeratedValue))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IEnumeratedValue) Implements System.Collections.Generic.IEnumerable(Of IEnumeratedValue).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IEnumeratedValue)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IEnumeratedValue)(Me.ToArray, GetType(EnumeratedValue), GetType(IEnumeratedValue))
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