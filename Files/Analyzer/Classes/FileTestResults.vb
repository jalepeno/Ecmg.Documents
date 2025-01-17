'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  FileTestResults.vb
'   Description :  [type_description_here]
'   Created     :  1/28/2015 1:38:06 PM
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

  Public Class FileTestResults
    Inherits List(Of FileTestResult)
    Implements IFileTestResults

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IFileTestResult)

#End Region

#Region "ITestResults Implementation"

    Public ReadOnly Property PrimaryResult As IFileTestResult Implements IFileTestResults.PrimaryResult
      Get
        Try
          If Count = 0 Then
            Return Nothing
          End If

          ' The items should have been sorted by the Analyzer
          Return Me.Item(0)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads Sub Add(item As IFileTestResult) Implements ICollection(Of IFileTestResult).Add
      Try
        If TypeOf (item) Is FileTestResult Then
          MyBase.Add(DirectCast(item, FileTestResult))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements ICollection(Of IFileTestResult).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IFileTestResult) As Boolean Implements ICollection(Of IFileTestResult).Contains
      Try
        Return Contains(CType(item, FileTestResult))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IFileTestResult, arrayIndex As Integer) Implements ICollection(Of IFileTestResult).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads ReadOnly Property Count As Integer Implements ICollection(Of IFileTestResult).Count
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

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of IFileTestResult).IsReadOnly
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

    Public Overloads Function Remove(item As IFileTestResult) As Boolean Implements ICollection(Of IFileTestResult).Remove
      Try
        Return MyBase.Remove(CType(item, FileTestResult))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As IEnumerator(Of IFileTestResult) Implements IEnumerable(Of IFileTestResult).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function IndexOf(item As IFileTestResult) As Integer Implements IList(Of IFileTestResult).IndexOf
      Try
        Return MyBase.IndexOf(CType(item, FileTestResult))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub Insert(index As Integer, item As IFileTestResult) Implements IList(Of IFileTestResult).Insert
      Try
        MyBase.Insert(index, CType(item, FileTestResult))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Default Public Overloads Property Item(index As Integer) As IFileTestResult Implements IList(Of IFileTestResult).Item
      Get
        Try
          Return MyBase.Item(index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IFileTestResult)
        Try
          MyBase.Item(index) = CType(value, FileTestResult)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overloads Sub RemoveAt(index As Integer) Implements IList(Of IFileTestResult).RemoveAt
      Try
        MyBase.RemoveAt(index)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IFileTestResult)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IFileTestResult)(Me.ToArray, GetType(FileTestResult), GetType(IFileTestResult))
          End If

          mobjEnumerator.OrderBy(Function(item) item.Percentage)

          Return mobjEnumerator
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#End Region

#Region "Private Methods"

    Friend Sub OrderByPercentage()
      Try
        Dim lobjQuery As Object = From results In Me Order By results.Percentage Descending Select results
        Dim lobjResults As New FileTestResults

        For Each lobjResult As FileTestResult In lobjQuery
          lobjResults.Add(lobjResult)
        Next

        Clear()

        For Each lobjResult As FileTestResult In lobjResults
          Add(lobjResult)
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace