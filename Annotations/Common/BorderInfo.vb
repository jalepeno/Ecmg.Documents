'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Annotations.Common
Imports Documents.Annotations.Decoration

#End Region

Namespace Annotations

  ''' <summary>
  ''' Specifies the properties of an annotation's border
  ''' </summary>
  <Serializable()>
  Public Class BorderInfo

#Region "Public Properties"

    ' ''' <summary>
    ' ''' Gets or sets the thickness.
    ' ''' </summary>
    ' ''' <value>The thickness.</value>
    '<XmlAttribute()> _
    'Public Property Thickness() As Integer

    ''' <summary>
    ''' Gets or sets the color.
    ''' </summary>
    ''' <value>The color.</value>
    Public Property Color() As ColorInfo

    ''' <summary>
    ''' Gets or sets the line style.
    ''' </summary>
    ''' <value>The line style.</value>
    Public Property LineStyle() As LineStyleInfo

#End Region

  End Class

End Namespace