'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Data
  Public Class OrderItems
    'Inherits Search.OrderItems
    Inherits Core.CCollection(Of OrderItem)

#Region "Overloads"

    Public Overloads Function Contains(ByVal fieldName As String) As Boolean
      For Each lobjOrderItem As OrderItem In Me
        If String.Equals(lobjOrderItem.FieldName, fieldName, StringComparison.InvariantCultureIgnoreCase) Then
          Return True
        End If
      Next
      Return False
    End Function

    Default Overridable Shadows Property Item(ByVal fieldName As String) As OrderItem
      Get
        For Each lobjOrderItem As OrderItem In Me
          If String.Equals(lobjOrderItem.FieldName, fieldName, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjOrderItem
          End If
        Next
        Return Nothing
      End Get
      Set(ByVal value As OrderItem)
        Dim lobjItem As Object
        ' Added to deal with strings in the ccollection
        If (value.GetType.Name = "String") Then
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjItem = CType(Item(lintCounter), OrderItem)
            If String.Equals(lobjItem, fieldName, StringComparison.InvariantCultureIgnoreCase) Then
              MyBase.Item(lintCounter) = value
              Exit Property
            End If
          Next
          Throw New Exception("There is no Item by the Name '" & fieldName & "'.")
        Else
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjItem = CType(Item(lintCounter), OrderItem)
            If String.Equals(lobjItem.Name, fieldName, StringComparison.InvariantCultureIgnoreCase) Then
              MyBase.Item(lintCounter) = value
              Exit Property
            End If
          Next
          Throw New Exception("There is no Item by the Name '" & fieldName & "'.")
        End If
      End Set
    End Property

    Default Overridable Shadows Property Item(ByVal index As Integer) As OrderItem
      Get
        Return MyBase.Item(index)
      End Get
      Set(ByVal value As OrderItem)
        MyBase.Item(index) = value
      End Set
    End Property

#End Region

  End Class

End Namespace