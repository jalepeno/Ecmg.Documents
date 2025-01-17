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

  ''' <summary>
  ''' Defines a reference to a document property
  ''' </summary>
  Public Class DocumentPropertyAnnotation
    Inherits ReferenceBase

#Region "Public Methods"

    ''' <summary>
    ''' Sets the reference.
    ''' </summary>
    ''' <param name="item">The item.</param>
    Public Overrides Sub SetReference(ByVal item As Object)
      If item Is Nothing Then
        Me.ReferenceContent = Nothing
      Else
        If TypeOf item Is ECMProperty Then
          Dim castItem As ECMProperty = DirectCast(item, ECMProperty)
          'Me.ReferenceContent = New ECMProperty(PropertyType.ecmString, "ReferenceContent", Cardinality.ecmSingleValued, castItem.Name, False)
          Me.ReferenceContent = PropertyFactory.Create(PropertyType.ecmString, "ReferenceContent", castItem.Name, Cardinality.ecmSingleValued)
        End If
      End If
    End Sub

#End Region

  End Class
End Namespace