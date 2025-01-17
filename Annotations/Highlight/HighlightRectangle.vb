'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Annotations.Highlight
  ''' <summary>
  ''' Indicates that this specific region is a rectangle, without further examination into its properties.
  ''' Provides greater portability for systems that only provide rectangular highlighted regions.
  ''' </summary>
  <Serializable()>
  Public Class HighlightRectangle
    Inherits HighlightBase

  End Class
End Namespace