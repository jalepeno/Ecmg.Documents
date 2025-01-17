'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core

#End Region

Namespace Annotations.Reference

  <Serializable()>
  <XmlInclude(GetType(DocumentPropertyAnnotation))>
  <XmlInclude(GetType(PageReferenceAnnotation))>
  <XmlInclude(GetType(UriReferenceAnnotation))>
  Public MustInherit Class ReferenceBase
    Inherits Annotation

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="ReferenceBase" /> class.
    ''' </summary>
    Protected Sub New()
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the content of the reference.
    ''' </summary>
    ''' <value>The content of the reference.</value>
    Public Property ReferenceContent() As ECMProperty

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Sets the reference.
    ''' </summary>
    ''' <param name="item">The item.</param>
    Public MustOverride Sub SetReference(ByVal item As Object)

#End Region

  End Class

End Namespace