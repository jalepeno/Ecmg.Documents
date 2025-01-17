'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Annotations.Reference

  Public Class PageReferenceAnnotation
    Inherits ReferenceBase

#Region "Public Methods"

    Public Overrides Sub SetReference(ByVal item As Object)
      If item Is Nothing Then
        Me.ReferenceContent = Nothing
      Else
        If TypeOf item Is Integer Then
          Dim castItem As Integer = DirectCast(item, Integer)
          'Me.ReferenceContent = New ECMProperty(PropertyType.ecmLong, "ReferenceContent", Cardinality.ecmSingleValued, castItem, False)
          Me.ReferenceContent = PropertyFactory.Create(PropertyType.ecmLong, "ReferenceContent", castItem)
        End If
      End If
    End Sub

#End Region

  End Class

End Namespace
