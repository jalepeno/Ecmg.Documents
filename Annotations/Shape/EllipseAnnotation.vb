'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Annotations.Common

#End Region

Namespace Annotations.Shape

  ''' <summary>
  ''' Defines a round-type annotation that is not properly modeled by a polygon.
  ''' </summary>
  <Serializable()>
  Public Class EllipseAnnotation
    Inherits ShapeBase

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the first focus.
    ''' </summary>
    ''' <value>The first focus.</value>
    Public Property FirstFocus() As Point

    ''' <summary>
    ''' Gets or sets the second focus.
    ''' </summary>
    ''' <value>The second focus.</value>
    Public Property SecondFocus() As Point

    ''' <summary>
    ''' Gets or sets the start angle.
    ''' </summary>
    ''' <value>The start angle.</value>
    <XmlAttribute()>
    Public Property StartAngle() As Single

    ''' <summary>
    ''' Gets or sets the end angle.
    ''' </summary>
    ''' <value>The end angle.</value>
    <XmlAttribute()>
    Public Property EndAngle() As Single

#End Region

    Public Sub SetInscribedEllipse()

    End Sub

  End Class

End Namespace