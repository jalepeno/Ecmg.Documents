'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization

#End Region

Namespace Annotations.Common

  ''' <summary>
  ''' Provides a representation of an arrow
  ''' </summary>
  <Serializable()>
  Public Class ArrowAnnotation
    Inherits LineBase

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the arrowhead size.
    ''' </summary>
    ''' <value>The arrowhead size.</value>
    <XmlAttribute()>
    Public Property Size As Integer

    ''' <summary>
    ''' Gets or sets the start point.
    ''' </summary>
    ''' <value>The start point.</value>
    Public Property StartPoint As Point

    ''' <summary>
    ''' Gets or sets the end point.
    ''' </summary>
    ''' <value>The end point.</value>
    Public Property EndPoint As Point

#End Region

  End Class

End Namespace