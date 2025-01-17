'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Annotations.Common
Imports Documents.Utilities

#End Region

Namespace Annotations.Shape

  <Serializable()>
  Public Class PolygonAnnotation
    Inherits ShapeBase

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="PolygonAnnotation" /> class.
    ''' </summary>
    Public Sub New()
      MyBase.New()
      Try
        Me.Points = New PointCollectionAnnotation()
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    Public Property Points() As PointCollectionAnnotation

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is a closed shape.
    ''' </summary>
    ''' <value><c>true</c> if this instance is closed; otherwise, <c>false</c>.</value>
    <XmlAttribute("Closed")>
    Public Property IsClosed() As Boolean

#End Region

  End Class

End Namespace