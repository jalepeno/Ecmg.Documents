'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization

#End Region

Namespace Annotations.Decoration

  ''' <summary>
  ''' Defines the thickness, shape and filled status of the point.
  ''' </summary>
  <Serializable()>
  Public Class PointStyle

#Region "Enumerations"

    Public Enum EndpointStyle
      None
      Dot
      Arrow
    End Enum

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets a value indicating whether this <see cref="PointStyle"/> is filled.
    ''' </summary>
    ''' <value><c>true</c> if filled; otherwise, <c>false</c>.</value>
    <XmlAttribute()>
    Public Property Filled() As Boolean

    ''' <summary>
    ''' Gets or sets the thickness.
    ''' </summary>
    ''' <value>The thickness.</value>
    <XmlAttribute()>
    Public Property Thickness() As Integer

    <XmlAttribute()>
    Public Property Endpoint As EndpointStyle

#End Region

  End Class

End Namespace
