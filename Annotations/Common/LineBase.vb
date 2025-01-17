'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"
Imports System.Xml.Serialization
Imports Documents.Annotations.Decoration

#End Region

Namespace Annotations.Common

  ''' <summary>
  ''' Marker class for potentially non-portable annotation types
  ''' </summary>
  <Serializable()>
  <XmlInclude(GetType(ArrowAnnotation))>
  <XmlInclude(GetType(PointCollectionAnnotation))>
  Public MustInherit Class LineBase
    Inherits Annotation

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the color.
    ''' </summary>
    ''' <value>The color.</value>
    Public Property Color As ColorInfo

    ''' <summary>
    ''' Gets or sets the line style.
    ''' </summary>
    ''' <value>The line style.</value>
    Public Property LineStyle As LineStyleInfo

    ''' <summary>
    ''' Gets or sets the width.
    ''' </summary>
    ''' <value>The width.</value>
    <XmlAttribute()>
    Public Property Width As Integer

#End Region

  End Class
End Namespace