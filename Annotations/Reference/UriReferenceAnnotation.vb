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

  Public Class UriReferenceAnnotation
    Inherits ReferenceBase

#Region "Public Methods"

    Public Overrides Sub SetReference(ByVal item As Object)
      If item Is Nothing Then
        Me.ReferenceContent = Nothing
      Else
        If TypeOf item Is Uri Then
          Dim castItem As Uri = DirectCast(item, Uri)
          'Me.ReferenceContent = New ECMProperty(PropertyType.ecmString, "ReferenceContent", Cardinality.ecmSingleValued, castItem.ToString, False)
          Me.ReferenceContent = PropertyFactory.Create(PropertyType.ecmString, "ReferenceContent", castItem.ToString)
        End If
      End If
    End Sub

#End Region

  End Class

End Namespace
