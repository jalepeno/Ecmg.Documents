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

Namespace Annotations.Text

  <Serializable()>
  Public Class FontInfo

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the color.
    ''' </summary>
    ''' <value>The color.</value>
    Public Property Color() As ColorInfo

    ''' <summary>
    ''' Gets or sets the name of the font.
    ''' </summary>
    ''' <value>The name of the font.</value>
    <XmlAttribute()>
    Public Property FontName() As String

    <XmlAttribute()>
    Public Property FontFamily() As String

    ''' <summary>
    ''' Gets or sets the size of the font.
    ''' </summary>
    ''' <value>The size of the font.</value>
    <XmlAttribute()>
    Public Property FontSize() As Single

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is bold.
    ''' </summary>
    ''' <value><c>true</c> if this instance is bold; otherwise, <c>false</c>.</value>
    <XmlAttribute()>
    Public Property IsBold() As Boolean

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is italic.
    ''' </summary>
    ''' <value>
    ''' <c>true</c> if this instance is italic; otherwise, <c>false</c>.
    ''' </value>
    <XmlAttribute()>
    Public Property IsItalic() As Boolean

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is strikethrough.
    ''' </summary>
    ''' <value>
    ''' <c>true</c> if this instance is strikethrough; otherwise, <c>false</c>.
    ''' </value>
    <XmlAttribute()>
    Public Property IsStrikethrough() As Boolean

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is underline.
    ''' </summary>
    ''' <value>
    ''' <c>true</c> if this instance is underline; otherwise, <c>false</c>.
    ''' </value>
    <XmlAttribute()>
    Public Property IsUnderline() As Boolean

#End Region

  End Class

End Namespace