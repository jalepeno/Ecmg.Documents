'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Annotations.Decoration
Imports Documents.Utilities

#End Region

Namespace Annotations.Shape

  <Serializable()>
  <XmlInclude(GetType(EllipseAnnotation))>
  <XmlInclude(GetType(RectangleAnnotation))>
  <XmlInclude(GetType(PolygonAnnotation))>
  Public MustInherit Class ShapeBase
    Inherits Annotation

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="EllipseAnnotation"/> class.
    ''' </summary>
    Protected Sub New()
      Try
        Me.LineStyle = New LineStyleInfo()
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the line style.
    ''' </summary>
    ''' <value>The line style.</value>
    Public Property LineStyle() As LineStyleInfo

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is filled.
    ''' </summary>
    ''' <value><c>true</c> if this instance is filled; otherwise, <c>false</c>.</value>
    ''' 
    <XmlAttribute("Filled")>
    Public Property IsFilled() As Boolean

#End Region

  End Class

End Namespace