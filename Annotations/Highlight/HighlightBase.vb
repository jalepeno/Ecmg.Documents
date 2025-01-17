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

Namespace Annotations.Highlight

  ''' <summary>
  ''' The base class for all annotations whose explicit purpose is to be highlighted regions.
  ''' </summary>
  <Serializable()>
  <XmlInclude(GetType(HighlightRectangle))>
  Public MustInherit Class HighlightBase
    Inherits Annotation

#Region "Public Properties"
    ''' <summary>
    ''' Gets or sets the color of the highlight.
    ''' </summary>
    ''' <value>The color of the highlight.</value>
    Public Property HighlightColor As ColorInfo
#End Region

#Region "Public Methods"
    Public Sub SetHighlght(ByVal redLevel As Integer, ByVal greenLevel As Integer, ByVal blueLevel As Integer)

      ' ----------------------------------------------------------------------------------
      ' Annotation display
      ' ----------------------------------------------------------------------------------
      Me.HighlightColor = New ColorInfo() With
        {
          .Red = redLevel,
          .Green = greenLevel,
          .Blue = blueLevel,
          .Opacity = 50
        }

    End Sub

#End Region

  End Class

End Namespace