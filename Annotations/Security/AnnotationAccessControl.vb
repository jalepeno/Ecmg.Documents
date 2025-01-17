'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization

#End Region

Namespace Annotations.Security

  ''' <summary>
  ''' Specifies the level of authoring permissions for an Annotation.
  ''' </summary>
  <Serializable()>
  Public Class AnnotationAccessControl

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the name of the principal user or goup name.
    ''' </summary>
    ''' <value>The name of the user or group.</value>
    <XmlAttribute()>
    Public Property PrincipalName() As String

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is group.
    ''' </summary>
    ''' <value><c>true</c> if this instance is group; otherwise, <c>false</c>.</value>
    <XmlAttribute()>
    Public Property IsGroup() As Boolean

    ''' <summary>
    ''' Gets or sets the permission assigned to the annotation.
    ''' </summary>
    ''' <value>The permission assigned to the annotation.</value>
    <XmlElement("AnnotationPermission")>
    Public Property Permission() As Permission

    ''' <summary>
    ''' Gets or sets the service provider.
    ''' </summary>
    ''' <value>The service provider.</value>
    <XmlAttribute()>
    Public Property ServiceProvider() As String

    ''' <summary>
    ''' Gets or sets the service context.
    ''' </summary>
    ''' <value>The service context.</value>
    <XmlAttribute()>
    Public Property ServiceContext() As String

#End Region

  End Class
End Namespace