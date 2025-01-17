'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class Repositories
    Inherits CCollection(Of Repository)
    Implements INamedItems

#Region "Public Methods"

    Public Shadows Sub Add(ByVal item As Repository)
      Try
        MyBase.Add(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ItemsByProductName(lpProductName As String) As Repositories
      Try

        Dim lobjReturnedItems As New Repositories

        Dim list As Object = From lobjRepository In Items Where
          lobjRepository.ProviderSystem.ProductName = lpProductName Select lobjRepository

        For Each lobjRepository As Repository In list
          lobjReturnedItems.Add(lobjRepository)
        Next

        Return lobjReturnedItems

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "INamedItems Implementation"

    Public Shadows Sub Add(ByVal item As Object) Implements INamedItems.Add
      Try
        If TypeOf (item) Is Repository Then
          Add(DirectCast(item, Repository))
        Else
          Throw New ArgumentException(
            String.Format("Invalid object type.  Item type of Repository expected instead of type {0}",
                          item.GetType.Name), "item")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Property ItemByName(ByVal name As String) As Object Implements INamedItems.ItemByName
      Get
        Return Item(name)
      End Get
      Set(ByVal value As Object)
        Item(name) = value
      End Set
    End Property

    Public Shadows Function Remove(ByVal lpItemName As String) As Boolean Implements INamedItems.Remove
      If Contains(lpItemName) Then
        Return MyBase.Remove(Me.Item(lpItemName))
      End If
    End Function

#End Region

  End Class

End Namespace
