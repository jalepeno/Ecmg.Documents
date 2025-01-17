'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Namespace Search

  Public Class OrderItem
    Inherits Data.OrderItem

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal FieldName As String, ByVal SortDirect As SortDirection)
      MyBase.New(FieldName, SortDirect)
    End Sub

  End Class
End Namespace