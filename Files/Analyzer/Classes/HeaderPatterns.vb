'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  HeaderPatterns.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 2:03:51 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"


Imports Documents.Utilities

#End Region

Namespace Files

  <Serializable()> Public Class HeaderPatterns
    Inherits List(Of HeaderPattern)
    Implements IHeaderPatterns
    Implements IComparable

#Region "Class Variables"

    <NonSerialized()> Private mobjEnumerator As IEnumeratorConverter(Of IHeaderPattern)

#End Region

    Public Overloads Sub Add(item As IHeaderPattern) Implements ICollection(Of IHeaderPattern).Add
      Try
        If TypeOf (item) Is HeaderPattern Then
          MyBase.Add(DirectCast(item, HeaderPattern))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements ICollection(Of IHeaderPattern).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IHeaderPattern) As Boolean Implements ICollection(Of IHeaderPattern).Contains
      Try
        Return Contains(CType(item, HeaderPattern))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IHeaderPattern, arrayIndex As Integer) Implements ICollection(Of IHeaderPattern).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads ReadOnly Property Count As Integer Implements ICollection(Of IHeaderPattern).Count
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

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of IHeaderPattern).IsReadOnly
      Get
        Try
          Return False
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads Function Remove(item As IHeaderPattern) As Boolean Implements ICollection(Of IHeaderPattern).Remove
      Try
        Return MyBase.Remove(CType(item, HeaderPattern))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As IEnumerator(Of IHeaderPattern) Implements IEnumerable(Of IHeaderPattern).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function IndexOf(item As IHeaderPattern) As Integer Implements IList(Of IHeaderPattern).IndexOf
      Try
        Return MyBase.IndexOf(CType(item, HeaderPattern))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub Insert(index As Integer, item As IHeaderPattern) Implements IList(Of IHeaderPattern).Insert
      Try
        MyBase.Insert(index, CType(item, HeaderPattern))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Default Public Overloads Property Item(index As Integer) As IHeaderPattern Implements IList(Of IHeaderPattern).Item
      Get
        Try
          Return MyBase.Item(index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IHeaderPattern)
        Try
          MyBase.Item(index) = CType(value, HeaderPattern)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads Sub RemoveAt(index As Integer) Implements IList(Of IHeaderPattern).RemoveAt
      Try
        MyBase.RemoveAt(index)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IHeaderPattern)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IHeaderPattern)(Me.ToArray, GetType(HeaderPattern), GetType(IHeaderPattern))
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

#Region "IComparable Implementation"

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try
        Dim lintItemComparisonScore As Integer = Me.Count.CompareTo(obj.Count)
        If lintItemComparisonScore <> 0 Then
          ' The counts do not match, we are not equal
          Return lintItemComparisonScore
        Else
          For lintPatternCounter As Integer = 0 To Me.Count - 1
            lintItemComparisonScore += DirectCast(Item(lintPatternCounter), HeaderPattern).CompareTo(DirectCast(obj.item(lintPatternCounter), HeaderPattern))
          Next
        End If

        Return lintItemComparisonScore

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace