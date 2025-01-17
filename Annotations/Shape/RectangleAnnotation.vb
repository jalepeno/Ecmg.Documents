'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Annotations.Common

#End Region

Namespace Annotations.Shape

  <Serializable()>
  Public Class RectangleAnnotation
    Inherits ShapeBase

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the upper left corner of the rectangle
    ''' </summary>
    ''' <value>The upper left corner.</value>
    Public Property UpperLeft() As Point

    ''' <summary>
    ''' Gets or sets the lower right corner of the rectangle
    ''' </summary>
    ''' <value>The lower right corner.</value>
    Public Property LowerRight() As Point

#End Region

  End Class
End Namespace
