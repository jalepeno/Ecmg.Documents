'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core

  ''' <summary>Collection of Relationship objects.</summary>
  Public Class Relationships
    Inherits CCollection(Of Relationship)

#Region "Public Properties"

    Public ReadOnly Property MaxOrder() As Integer
      Get
        Dim lintMax As Integer
        For Each lobjRelationship As Relationship In Me
          If lobjRelationship.Order > lintMax Then
            lintMax = lobjRelationship.Order
          End If
        Next
        Return lintMax
      End Get
    End Property

#End Region

#Region "Overloads"

    Default Shadows Property Item(ByVal ObjectID As String) As Relationship
      Get
        Dim lobjVersion As Relationship
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjVersion = CType(MyBase.Item(lintCounter), Relationship)
          If lobjVersion.ObjectID = ObjectID Then
            Return lobjVersion
          End If
        Next
        Throw New Exception("There is no Relationship by the ObjectID '" & ObjectID & "'.")
        'Throw New InvalidArgumentException
      End Get
      Set(ByVal value As Relationship)
        Dim lobjVersion As Relationship
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjVersion = CType(MyBase.Item(lintCounter), Relationship)
          If lobjVersion.ObjectID = ObjectID Then
            MyBase.Item(lintCounter) = value
          End If
        Next
        Throw New Exception("There is no Relationship by the ObjectID '" & ObjectID & "'.")
      End Set
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As Relationship
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As Relationship)
        MyBase.Item(Index) = value
      End Set
    End Property

    Public Function Delete(ByVal ObjectID As String) As Boolean
      Try
        Remove(Item(ObjectID))
        Return True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try
    End Function

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Verfies that all of the relationships in the collection are valid.
    ''' </summary>
    ''' <returns>True or False</returns>
    ''' <remarks>If there are no relationships then False is returned</remarks>
    Public Function AllValid(ByVal lpParentDocument As Document) As Boolean
      Try
        If InvalidRelationshipCount(lpParentDocument) = 0 Then
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

    Public Function InvalidRelationshipCount(ByVal lpParentDocument As Document) As Integer
      Try
        If Count = 0 Then
          Return 0
        End If

        Dim lintInvalidRelationshipCounter As Integer
        For Each lobjRelationship As Relationship In Me
          If lobjRelationship.IsValid(lpParentDocument) = False Then
            lintInvalidRelationshipCounter += 1
          End If
        Next

        Return lintInvalidRelationshipCounter

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace