'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Data
  Public Class OrderItem
    'Inherits Search.OrderItem

    Inherits Core.Value

#Region "Class Variables"
    Private mstrFieldName As String = ""
    Private mstrSortDirection As SortDirection
#End Region

#Region "Public Properties"

    Public Property FieldName() As String
      Get
        Return mstrFieldName
      End Get
      Set(ByVal value As String)
        mstrFieldName = value
      End Set
    End Property

    Public Property SortDirection() As SortDirection
      Get
        Return mstrSortDirection
      End Get
      Set(ByVal value As SortDirection)
        mstrSortDirection = value
      End Set
    End Property

#End Region

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal FieldName As String, ByVal SortDirect As SortDirection)
      MyBase.New()
      Me.mstrFieldName = FieldName
      Me.mstrSortDirection = SortDirect
    End Sub

  End Class
End Namespace