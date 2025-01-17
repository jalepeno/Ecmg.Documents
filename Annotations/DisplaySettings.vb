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

Namespace Annotations

  ''' <summary>
  ''' Defines the color, transparency, border and overall visibility of an annotation.
  ''' </summary>
  <Serializable()>
  Public Class DisplaySettings

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="DisplaySettings"/> class.
    ''' </summary>
    Public Sub New()
      Me.Visible = True
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets a value indicating whether this <see cref="DisplaySettings"/> is visible.
    ''' </summary>
    ''' <value><c>true</c> if visible; otherwise, <c>false</c>.</value>
    <XmlAttribute()>
    Public Property Visible() As Boolean

    ''' <summary>
    ''' Gets or sets the foreground.
    ''' </summary>
    ''' <value>The foreground.</value>
    Public Property Foreground() As ColorInfo

    ''' <summary>
    ''' Gets or sets the background.
    ''' </summary>
    ''' <value>The background.</value>
    Public Property Background() As ColorInfo

    ''' <summary>
    ''' Gets or sets the border.
    ''' </summary>
    ''' <value>The border.</value>
    Public Property Border() As BorderInfo

#End Region

  End Class

End Namespace