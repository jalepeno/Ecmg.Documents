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
  ''' Provides the ability to specify a 24-bit color and opacity level (alpha)
  ''' </summary>
  <Serializable()>
  Public Class ColorInfo

#Region "Public Properties"

    <XmlAttribute("R")>
    Public Property Red As Integer

    <XmlAttribute("G")>
    Public Property Green As Integer

    <XmlAttribute("B")>
    Public Property Blue As Integer

    ''' <summary>
    ''' Gets or sets the opacity.
    ''' </summary>
    ''' <value>The opacity.</value>
    <XmlAttribute("A")>
    Public Property Opacity() As Integer

#End Region

  End Class

End Namespace