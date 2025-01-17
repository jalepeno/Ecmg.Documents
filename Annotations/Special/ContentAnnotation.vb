'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Annotations.Special

  <Serializable()>
  Public Class ContentAnnotation
    Inherits SpecialBase

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the content element.
    ''' </summary>
    ''' <value>The content element.</value>
    Public Property ContentElement() As Content

#End Region

  End Class

End Namespace