'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Annotations.Decoration

  ''' <summary>
  ''' Defines thickness, pattern and path type for the line.
  ''' </summary>
  <Serializable()>
  Public Class LineStyleInfo

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="LineStyleInfo"/> class.
    ''' </summary>
    Public Sub New()
      Try
        'Me.Layout = New LayoutSettings()
        Me.LineWeight = 1
        Me.Pattern = LinePattern.Solid
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Enumerations"

    Public Enum LinePattern
      None
      Solid
      Dot
      Dash
      DashDot
      DashDotDot
    End Enum

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the pattern.
    ''' </summary>
    ''' <value>The pattern.</value>
    Public Property Pattern() As LinePattern

    ''' <summary>
    ''' Gets or sets the line weight.
    ''' </summary>
    ''' <value>The line weight.</value>
    <XmlAttribute()>
    Public Property LineWeight() As Integer

#End Region

  End Class
End Namespace